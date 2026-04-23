"""
bulk_run.py — Process all client holdings files and output a master transaction CSV.

SUPPORTS TWO INPUT FORMATS:
  Format A: One Excel file with one sheet per client
            Sheet name = client_id
            File: data/dummy_clients/all_clients.xlsx (or any xlsx with multiple sheets)

  Format B: One Excel/CSV file per client in a directory
            Filename (without extension) = client_id
            Dir: data/dummy_clients/by_client/

SMART COLUMN DETECTION:
  Does not rely on fixed column names. Instead:
  - ISIN column: find the column where ≥50% of non-null values match
                 the Indian ISIN pattern (INF/INE + 10 chars, 12 total)
  - Units column: among numeric columns, prefer the one whose header
                  contains 'unit', 'qty', 'quantity', 'balance', 'holding'
                  Falls back to the numeric column most correlated with
                  a typical units range (positive, non-zero)

Usage:
  python bulk_run.py --format A --file data/dummy_clients/all_clients.xlsx
  python bulk_run.py --format B --dir  data/dummy_clients/by_client/
  python bulk_run.py --format B        (uses default CLIENT_DIR from config)
"""

import argparse
import os
import re
import sys

import pandas as pd

from config import (
    ARCHETYPES,
    CLIENT_AGES_FILE,
    CLIENT_DIR,
    FUND_CHOICES_FILE,
    OUTPUT_CSV,
)
from fund_choices import load_fund_choices, validate_fund_choices
from portfolio import compute_portfolio, compute_glide_target_mix, generate_transactions


CLIENT_AGES_SHEET = "ClientAges"


def display_risk_type(archetype_label: str) -> str:
    """
    Risk label for CSV/Excel export: Averse | Moderate | Aggressive.

    Maps internal glide labels like \"Glide:Moderate\" to \"Moderate\".
    """
    s = str(archetype_label).strip()
    if s.lower().startswith("glide:"):
        rest = s.split(":", 1)[1].strip()
        for k in ARCHETYPES:
            if k.lower() == rest.lower():
                return k
        return rest.title() if rest else "Moderate"
    for k in ARCHETYPES:
        if k.lower() == s.lower():
            return k
    return s if s else "Moderate"


def load_client_ages(path: str) -> dict[str, dict]:
    """
    Load sidecar workbook with sheet ClientAges.
    Columns: client_id, age, risk_preference (header names case/spacing tolerant).
    Returns map client_id -> {"age": int/float, "risk_preference": str}.
    """
    df = pd.read_excel(path, sheet_name=CLIENT_AGES_SHEET)
    norm = {str(c).strip().lower().replace(" ", "_"): c for c in df.columns}
    def pick(*names):
        for n in names:
            if n in norm:
                return norm[n]
        return None

    c_id = pick("client_id", "clientid")
    age_c = pick("age")
    risk_c = pick("risk_preference", "riskpreference", "preference")
    if not c_id or not age_c or not risk_c:
        raise ValueError(
            f"{path}: sheet {CLIENT_AGES_SHEET!r} needs columns "
            "client_id, age, risk_preference "
            f"(got {list(df.columns)})"
        )
    out = {}
    for _, row in df.iterrows():
        cid = str(row[c_id]).strip()
        if not cid or cid.lower() == "nan":
            continue
        out[cid] = {
            "age": row[age_c],
            "risk_preference": str(row[risk_c]).strip(),
        }
    return out


def _resolve_new_scheme_pools(
    allow_new_funds: bool,
    fund_choices_path: str | None,
    master_df: pd.DataFrame,
) -> dict[str, list[tuple[str, float]]] | None:
    if not allow_new_funds:
        return None
    path = fund_choices_path or FUND_CHOICES_FILE
    if not os.path.isfile(path):
        ap = os.path.abspath(path)
        raise FileNotFoundError(
            "allow_new_funds requires a FundChoices workbook, but this file "
            f"does not exist:\n  {ap}\n"
            "Create or refresh it from your NAV master (same workbook as the pack uses):\n"
            f"  python fund_choices.py --master <path\\to\\latestNAV_Reports.xlsx> "
            f"--out {path}"
        )
    pools = load_fund_choices(path)
    nrows = sum(len(v) for v in pools.values())
    if nrows == 0:
        print(
            "  WARNING: FundChoices empty — buys use existing holdings only",
            file=sys.stderr,
        )
        return None
    bad = validate_fund_choices(pools, master_df)
    if bad:
        for m in bad:
            print(f"  ERROR fund_choices: {m}", file=sys.stderr)
        raise ValueError("fund_choices validation failed — fix workbook or master U–X")
    print(
        f"  New-fund pools: {path} | "
        f"equity={len(pools['equity'])} defensive={len(pools['defensive'])} "
        f"other={len(pools['other'])}"
    )
    return pools


ISIN_PATTERN = re.compile(r'^IN[A-Z0-9]{10}$')


# ── Column detection ──────────────────────────────────────────────────────────

def detect_isin_column(df: pd.DataFrame) -> str | None:
    """
    Find the column whose values look like Indian ISINs.
    Returns column name or None if not found.
    """
    for col in df.columns:
        sample = df[col].dropna().astype(str).str.strip()
        if len(sample) == 0:
            continue
        match_rate = sample.apply(lambda v: bool(ISIN_PATTERN.match(v))).mean()
        if match_rate >= 0.5:
            return col
    return None


def detect_units_column(df: pd.DataFrame, isin_col: str) -> str | None:
    """
    Find the units/quantity column.
    Preference order:
      1. Numeric column whose header contains 'unit', 'qty', 'quantity',
         'balance', 'holding'
      2. Any numeric column with all-positive values and reasonable range
         (units are typically 0.01 – 1,000,000)
    """
    unit_keywords = ["unit", "qty", "quantity", "balance", "holding"]

    numeric_cols = [
        c for c in df.columns
        if c != isin_col and pd.api.types.is_numeric_dtype(df[c])
    ]

    # Priority 1: keyword match in header
    for col in numeric_cols:
        if any(kw in col.lower() for kw in unit_keywords):
            return col

    # Priority 2: first numeric col with all-positive values
    for col in numeric_cols:
        vals = df[col].dropna()
        if len(vals) > 0 and (vals > 0).all():
            return col

    return None


def parse_holdings(df: pd.DataFrame, source_label: str) -> pd.DataFrame | None:
    """
    Given a raw DataFrame, detect ISIN and Units columns and return
    a clean DataFrame with columns ['isin', 'units'].
    Returns None and prints a warning if detection fails.
    """
    isin_col  = detect_isin_column(df)
    if not isin_col:
        print(f"    SKIP {source_label}: could not detect ISIN column")
        return None

    units_col = detect_units_column(df, isin_col)
    if not units_col:
        print(f"    SKIP {source_label}: could not detect Units column "
              f"(ISIN col found: {isin_col!r})")
        return None

    result = df[[isin_col, units_col]].copy()
    result.columns = ["isin", "units"]
    result["isin"]  = result["isin"].astype(str).str.strip()
    result["units"] = pd.to_numeric(result["units"], errors="coerce")
    result = result[
        result["isin"].str.match(r'^IN[A-Z0-9]{10}$') &
        result["units"].notna() &
        (result["units"] > 0)
    ].reset_index(drop=True)

    if result.empty:
        print(f"    SKIP {source_label}: no valid ISIN+Units rows after filtering")
        return None

    return result


# ── Format A: multi-sheet Excel ───────────────────────────────────────────────

def process_format_a(
    file_path: str,
    master_df: pd.DataFrame,
    archetype: str,
    new_cash: float,
    client_ages: dict[str, dict] | None = None,
    new_scheme_pools: dict[str, list[tuple[str, float]]] | None = None,
) -> list[dict]:
    """Process a multi-sheet Excel file (one sheet = one client)."""
    xl = pd.ExcelFile(file_path)
    all_txns = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=0)
        holdings = parse_holdings(df, f"{os.path.basename(file_path)}[{sheet_name}]")
        if holdings is None:
            continue

        mix, label = _targets_for_client(sheet_name, archetype, client_ages)
        txns, summary = _run_client(
            client_id=sheet_name,
            holdings=holdings,
            master_df=master_df,
            archetype=label,
            new_cash=new_cash,
            target_mix=mix,
            new_scheme_pools=new_scheme_pools,
        )
        all_txns.extend(txns)

    return all_txns


# ── Format B: one file per client ────────────────────────────────────────────

def process_format_b(
    client_dir: str,
    master_df: pd.DataFrame,
    archetype: str,
    new_cash: float,
    client_ages: dict[str, dict] | None = None,
    new_scheme_pools: dict[str, list[tuple[str, float]]] | None = None,
) -> list[dict]:
    """Process a directory of client files (filename = client_id)."""
    if not os.path.isdir(client_dir):
        print(f"  [bulk_run] Directory not found: {client_dir}")
        return []

    files = [
        f for f in sorted(os.listdir(client_dir))
        if f.lower().endswith((".xlsx", ".xls", ".csv"))
    ]
    print(f"  Found {len(files)} client files in {client_dir}")
    all_txns = []

    for fname in files:
        client_id = os.path.splitext(fname)[0]
        if client_id.lower() in ("client_ages",):
            continue
        fpath     = os.path.join(client_dir, fname)
        df = (pd.read_csv(fpath) if fname.lower().endswith(".csv")
              else pd.read_excel(fpath))

        holdings = parse_holdings(df, fname)
        if holdings is None:
            continue

        mix, label = _targets_for_client(client_id, archetype, client_ages)
        txns, summary = _run_client(
            client_id=client_id,
            holdings=holdings,
            master_df=master_df,
            archetype=label,
            new_cash=new_cash,
            target_mix=mix,
            new_scheme_pools=new_scheme_pools,
        )
        all_txns.extend(txns)

    return all_txns


# ── Core per-client computation ───────────────────────────────────────────────

def _targets_for_client(
    client_id: str,
    fallback_archetype: str,
    client_ages: dict[str, dict] | None,
) -> tuple[dict | None, str]:
    """
    If client_ages has this client, return glide target_mix and label \"Glide:…\".
    Otherwise return (None, fallback_archetype) for classic ARCHETYPES row.
    """
    if not client_ages or client_id not in client_ages:
        if client_ages is not None and client_id not in client_ages:
            print(
                f"    [age-based] {client_id}: no row in client ages — "
                f"using archetype {fallback_archetype!r}"
            )
        return None, fallback_archetype
    row = client_ages[client_id]
    age = row["age"]
    try:
        age_f = float(age)
    except (TypeError, ValueError):
        print(f"    [age-based] {client_id}: bad age {age!r} — using archetype")
        return None, fallback_archetype
    pref = str(row["risk_preference"])
    try:
        mix = compute_glide_target_mix(age_f, pref)
    except KeyError as e:
        print(f"    [age-based] {client_id}: {e} — using archetype")
        return None, fallback_archetype
    label = f"Glide:{pref.strip().title()}"
    return mix, label


def _run_client(
    client_id: str,
    holdings: pd.DataFrame,
    master_df: pd.DataFrame,
    archetype: str,
    new_cash: float,
    target_mix: dict | None = None,
    new_scheme_pools: dict[str, list[tuple[str, float]]] | None = None,
) -> tuple[list[dict], dict]:
    """Run portfolio computation and transaction generation for one client."""
    portfolio       = compute_portfolio(holdings, master_df)
    txns, summary   = generate_transactions(
        portfolio,
        archetype,
        new_cash=new_cash,
        target_mix=target_mix,
        new_scheme_pools=new_scheme_pools,
        master_df=master_df,
    )

    audit = {
        "new_cash_inr": round(float(summary["new_cash"]), 2),
        "deployed_new_cash_to_others_inr": float(
            summary["deployed_new_cash_to_others_inr"]
        ),
        "total_buy_inr": float(summary["total_buy_inr"]),
        "total_sell_inr": float(summary["total_sell_inr"]),
        "net_flow_inr": float(summary["net_flow_inr"]),
    }
    enriched = []
    for tx in txns:
        row = {
            "client_id":               client_id,
            "isin":                    tx["isin"],
            "scheme_name":             tx["scheme_name"],
            "action":                  tx["action"],
            "amount_inr":              tx["amount_inr"],
            "target_policy":          display_risk_type(summary["archetype"]),
            "current_equity_pct":      f"{summary['current_equity_pct']:.1%}",
            "target_equity_pct":       f"{summary['target_equity_pct']:.1%}",
            "current_defensive_pct":   f"{summary['current_defensive_pct']:.1%}",
            "target_defensive_pct":    f"{summary['target_defensive_pct']:.1%}",
            "current_other_pct":       f"{summary['current_other_pct']:.1%}",
            "target_other_pct":        f"{summary['target_other_pct']:.1%}",
            "portfolio_value_inr":     summary["total_portfolio"],
            "switch_amount_inr":       summary["switch_amount"],
        }
        row.update(audit)
        enriched.append(row)

    nf = audit["net_flow_inr"]
    nc = audit["new_cash_inr"]
    flow_warn = ""
    if abs(nf - nc) > 0.02:
        flow_warn = f" | WARN net_flow {nf:,.0f} != new_cash {nc:,.0f}"

    print(
        f"  {client_id}: INR {portfolio['total_value']:,.0f} | "
        f"eq={portfolio['equity_pct']:.1%} def={portfolio['defensive_pct']:.1%} "
        f"oth={portfolio['other_pct']:.1%} | "
        f"target={summary['target_equity_pct']:.0%}/"
        f"{summary['target_defensive_pct']:.0%}/"
        f"{summary['target_other_pct']:.0%} | "
        f"{len(txns)} tx | buy {audit['total_buy_inr']:,.0f} sell {audit['total_sell_inr']:,.0f} "
        f"net_flow {nf:,.0f}{flow_warn}"
    )
    return enriched, summary


# ── Main ──────────────────────────────────────────────────────────────────────

def detect_format(path: str) -> tuple[str, str | None, str | None]:
    """
    Auto-detect input format from path.

    Rules:
      - Directory              → Format B, client_dir = path
      - Excel with >1 sheets  → Format A, file_path  = path
      - Excel with 1 sheet    → Format B single file (dir = parent dir)
      - CSV                   → Format B single file

    Returns (format, file_path, client_dir)
    """
    if os.path.isdir(path):
        print(f"  Detected Format B (directory of client files): {path}")
        return "B", None, path

    if path.lower().endswith((".xlsx", ".xls")):
        xl = pd.ExcelFile(path)
        n_sheets = len(xl.sheet_names)
        if n_sheets > 1:
            print(f"  Detected Format A (multi-sheet Excel, {n_sheets} clients): {path}")
            return "A", path, None
        else:
            parent = os.path.dirname(path) or "."
            print(f"  Detected Format B single file (1-sheet Excel): {path}")
            return "B", None, parent

    if path.lower().endswith(".csv"):
        parent = os.path.dirname(path) or "."
        print(f"  Detected Format B single file (CSV): {path}")
        return "B", None, parent

    raise ValueError(f"Cannot detect format from path: {path!r}")


def bulk_run(
    format: str | None = None,
    file_path: str | None = None,
    client_dir: str = CLIENT_DIR,
    output_csv: str = OUTPUT_CSV,
    archetype: str = "Moderate",
    new_cash: float = 0,
    master_df: pd.DataFrame | None = None,
    path: str | None = None,
    age_based: bool = False,
    client_ages_path: str | None = None,
    allow_new_funds: bool = False,
    fund_choices_path: str | None = None,
) -> pd.DataFrame:
    """
    Run bulk rebalancing for all clients.

    Provide either:
      path   — auto-detects format (directory → B, multi-sheet Excel → A)
      format + file_path/client_dir — explicit format

    master_df : pre-loaded master DataFrame (loaded from config if not provided)
    """
    if master_df is None:
        from portfolio import load_master_for_portfolio
        from config import ISIN_MASTER_FILE
        master_df = load_master_for_portfolio(ISIN_MASTER_FILE)

    client_ages: dict[str, dict] | None = None
    if age_based:
        cap = client_ages_path or CLIENT_AGES_FILE
        if not os.path.isfile(cap):
            raise FileNotFoundError(
                f"Age-based mode requires client ages workbook: {cap}"
            )
        client_ages = load_client_ages(cap)
        print(
            f"\n  Age-based glide path ON | sidecar: {cap} ({len(client_ages)} rows) | "
            f"fallback archetype: {archetype!r} | new cash (INR): {new_cash:,.0f}"
        )
    else:
        print(
            f"\n  Archetype (fixed): {archetype} | new cash (INR): {new_cash:,.0f}"
        )

    new_scheme_pools = _resolve_new_scheme_pools(
        allow_new_funds, fund_choices_path, master_df
    )

    # Auto-detect if path given
    if path is not None:
        format, file_path, client_dir = detect_format(path)

    if format is not None and format.upper() == "A":
        if not file_path:
            raise ValueError("Format A requires a file path.")
        all_txns = process_format_a(
            file_path,
            master_df,
            archetype,
            new_cash,
            client_ages=client_ages,
            new_scheme_pools=new_scheme_pools,
        )
    else:
        all_txns = process_format_b(
            client_dir,
            master_df,
            archetype,
            new_cash,
            client_ages=client_ages,
            new_scheme_pools=new_scheme_pools,
        )

    if not all_txns:
        print("  No transactions generated.")
        return pd.DataFrame()

    df_out = pd.DataFrame(all_txns)
    os.makedirs(os.path.dirname(output_csv) or ".", exist_ok=True)
    df_out.to_csv(output_csv, index=False)
    print(f"\n  OK {len(all_txns)} transactions -> {output_csv}")
    return df_out


def main():
    parser = argparse.ArgumentParser(
        description="Bulk rebalancing for all client portfolios."
    )
    # Single --path argument auto-detects format
    parser.add_argument("--path",      default=None,
                        help="Directory (Format B) or multi-sheet Excel (Format A) — auto-detected")
    # Explicit overrides still available
    parser.add_argument("--format",    choices=["A", "B"], default=None,
                        help="Override auto-detection: A=multi-sheet Excel, B=one file per client")
    parser.add_argument("--file",      default=None,
                        help="Explicit Format A: path to multi-sheet Excel")
    parser.add_argument("--dir",       default=None,
                        help="Explicit Format B: client directory")
    parser.add_argument("--archetype", default="Moderate",
                        choices=list(ARCHETYPES.keys()))
    parser.add_argument("--new-cash",  type=float, default=0,
                        help="Fresh cash to deploy per client (INR)")
    parser.add_argument("--output",    default=OUTPUT_CSV,
                        help="Output CSV path")
    parser.add_argument(
        "--age-based",
        action="store_true",
        help="Use client_ages.xlsx glide path (100 - k×age); needs --client-ages or default path",
    )
    parser.add_argument(
        "--client-ages",
        default=None,
        dest="client_ages",
        help=f"Path to workbook with sheet ClientAges (default: {CLIENT_AGES_FILE})",
    )
    parser.add_argument(
        "--allow-new-funds",
        action="store_true",
        default=False,
        help="Allow BUYs into MFD-picked funds (sheet FundChoices in --fund-choices workbook)",
    )
    parser.add_argument(
        "--fund-choices",
        default=None,
        dest="fund_choices",
        help=f"Workbook with FundChoices (+ optional NAV copy); default: {FUND_CHOICES_FILE}",
    )
    args = parser.parse_args()

    from portfolio import load_master_for_portfolio
    from config import ISIN_MASTER_FILE

    print("\n=== Bulk Rebalancing Run ===")
    master_df = load_master_for_portfolio(ISIN_MASTER_FILE)
    print(f"  Master loaded: {len(master_df):,} ISINs")

    # Resolve path vs explicit format args
    if args.path:
        bulk_run(
            path              = args.path,
            output_csv        = args.output,
            archetype         = args.archetype,
            new_cash          = args.new_cash,
            master_df         = master_df,
            age_based         = args.age_based,
            client_ages_path  = args.client_ages,
            allow_new_funds   = args.allow_new_funds,
            fund_choices_path = args.fund_choices,
        )
    else:
        bulk_run(
            format            = args.format or "B",
            file_path         = args.file,
            client_dir        = args.dir or CLIENT_DIR,
            output_csv        = args.output,
            archetype         = args.archetype,
            new_cash          = args.new_cash,
            master_df         = master_df,
            age_based         = args.age_based,
            client_ages_path  = args.client_ages,
            allow_new_funds   = args.allow_new_funds,
            fund_choices_path = args.fund_choices,
        )
    print("\n=== Done ===\n")


if __name__ == "__main__":
    main()