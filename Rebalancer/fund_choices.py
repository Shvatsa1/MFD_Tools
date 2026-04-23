"""
fund_choices.py — Load / validate MFD-selected new funds for rebalancing.

Workbook sheet **FundChoices** (long format):

  sleeve   | isin        | fund_name (optional, for MFD) | weight
  -------- | ----------- | ------------------------------ | ------
  equity   | INE0...     | Axis Bluechip Fund            | 0.5

``sleeve`` is one of: equity, defensive, debt, other (debt/defensive treated the same).
``fund_name`` is optional at load time (ISIN is authoritative); templates populate it from col D.

Weights are relative (need not sum to 1); non-positive rows skipped.

Eligibility on master Final columns (U–X): see config NEW_SCHEME_MIN_*.
"""

from __future__ import annotations

import os

import pandas as pd

from config import (
    NEW_SCHEME_MIN_DEFENSIVE_FRAC,
    NEW_SCHEME_MIN_EQUITY_FRAC,
    NEW_SCHEME_MIN_OTHER_FRAC,
)
from portfolio import COL_FINAL_CASH, COL_FINAL_DEBT, COL_FINAL_EQUITY, COL_FINAL_OTHER

FUND_CHOICES_SHEET = "FundChoices"

NewSchemePools = dict[str, list[tuple[str, float]]]


def _norm_col(name: str) -> str:
    return str(name).strip().lower().replace(" ", "_")


def load_fund_choices(path: str) -> NewSchemePools:
    df = pd.read_excel(path, sheet_name=FUND_CHOICES_SHEET)
    colmap = {_norm_col(c): c for c in df.columns}
    sc = colmap.get("sleeve") or colmap.get("category") or colmap.get("bucket")
    ic = colmap.get("isin")
    wc = colmap.get("weight") or colmap.get("wt") or colmap.get("weight_pct")
    if not sc or not ic:
        raise ValueError(
            f"{path}: sheet {FUND_CHOICES_SHEET!r} needs sleeve (or category) and isin "
            f"(columns: {list(df.columns)})"
        )
    pools: NewSchemePools = {"equity": [], "defensive": [], "other": []}
    for _, row in df.iterrows():
        sleeve = str(row[sc]).strip().lower()
        if sleeve in ("debt", "defensive", "fixed_income", "def"):
            sleeve = "defensive"
        elif sleeve in ("equity", "stock", "eq"):
            sleeve = "equity"
        elif sleeve in ("other", "gold", "commodity"):
            sleeve = "other"
        if sleeve not in pools:
            continue
        isin = str(row[ic]).strip().upper()
        if not isin or isin == "NAN":
            continue
        wt = 1.0
        if wc and pd.notna(row[wc]):
            try:
                wt = float(row[wc])
            except (TypeError, ValueError):
                wt = 1.0
        if wt <= 0:
            continue
        pools[sleeve].append((isin, wt))

    return pools


def _row_eligible(row: pd.Series, side: str) -> bool:
    fe = float(pd.to_numeric(row.get(COL_FINAL_EQUITY), errors="coerce") or 0)
    fd = float(pd.to_numeric(row.get(COL_FINAL_DEBT), errors="coerce") or 0)
    fc = float(pd.to_numeric(row.get(COL_FINAL_CASH), errors="coerce") or 0)
    fo = float(pd.to_numeric(row.get(COL_FINAL_OTHER), errors="coerce") or 0)
    if side == "equity":
        return fe >= NEW_SCHEME_MIN_EQUITY_FRAC - 1e-9
    if side == "defensive":
        return (fd + fc) >= NEW_SCHEME_MIN_DEFENSIVE_FRAC - 1e-9
    if side == "other":
        return fo >= NEW_SCHEME_MIN_OTHER_FRAC - 1e-9
    return False


def validate_fund_choices(pools: NewSchemePools, master_df: pd.DataFrame) -> list[str]:
    """Return list of warning/error strings; empty if OK."""
    msgs: list[str] = []
    idx = master_df.set_index("isin", drop=False)
    for side, entries in pools.items():
        for isin, wt in entries:
            isin = str(isin).strip()
            if isin not in idx.index:
                msgs.append(f"{side}: ISIN not in master {isin!r}")
                continue
            row = idx.loc[isin]
            if isinstance(row, pd.DataFrame):
                row = row.iloc[0]
            if not _row_eligible(row, side):
                msgs.append(
                    f"{side}: ISIN {isin} fails sleeve purity "
                    f"(need equity≥{NEW_SCHEME_MIN_EQUITY_FRAC}, "
                    f"(debt+cash)≥{NEW_SCHEME_MIN_DEFENSIVE_FRAC} for defensive, "
                    f"other≥{NEW_SCHEME_MIN_OTHER_FRAC})"
                )
    return msgs


def build_fund_choices_template(master_path: str, out_path: str) -> NewSchemePools:
    """
    Build FundChoices + optional latestNAV_Reports copy for dropdown scaffolding.
    Picks up to 3 ISINs per sleeve that pass eligibility (stable sort by ISIN).
    """
    df = pd.read_excel(master_path, sheet_name="latestNAV_Reports")
    raw = list(df.columns)
    col_d, col_e = raw[3], raw[4]
    col_u, col_v, col_w, col_x = raw[20], raw[21], raw[22], raw[23]
    slim = df[[col_d, col_e, col_u, col_v, col_w, col_x]].copy()
    slim.columns = [
        "scheme_name",
        "isin",
        COL_FINAL_EQUITY,
        COL_FINAL_DEBT,
        COL_FINAL_CASH,
        COL_FINAL_OTHER,
    ]
    slim["isin"] = slim["isin"].astype(str).str.strip()
    slim["scheme_name"] = slim["scheme_name"].astype(str).str.strip()
    for c in (
        COL_FINAL_EQUITY,
        COL_FINAL_DEBT,
        COL_FINAL_CASH,
        COL_FINAL_OTHER,
    ):
        slim[c] = pd.to_numeric(slim[c], errors="coerce")

    cands = slim.drop_duplicates("isin").sort_values("isin")
    pools: NewSchemePools = {"equity": [], "defensive": [], "other": []}
    isin_fund_name: dict[str, str] = {}
    for side in ("equity", "defensive", "other"):
        for _, row in cands.iterrows():
            if len(pools[side]) >= 3:
                break
            isin = row["isin"]
            if not isin or str(isin).lower() == "nan" or len(str(isin)) < 10:
                continue
            if _row_eligible(row, side):
                pools[side].append((isin, 1.0))
                nm = row["scheme_name"]
                isin_fund_name[str(isin)] = (
                    "" if nm is None or str(nm).lower() in ("nan", "none") else str(nm).strip()
                )

    rows = []
    for sleeve, pairs in pools.items():
        for isin, wt in pairs:
            isin_s = str(isin).strip()
            rows.append({
                "sleeve": sleeve,
                "isin": isin_s,
                "fund_name": isin_fund_name.get(isin_s, ""),
                "weight": wt,
            })

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    fc = pd.DataFrame(rows, columns=["sleeve", "isin", "fund_name", "weight"])
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        fc.to_excel(w, sheet_name=FUND_CHOICES_SHEET, index=False)
        nav = pd.read_excel(master_path, sheet_name="latestNAV_Reports")
        nav.to_excel(w, sheet_name="latestNAV_Reports", index=False)
    return pools


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Build fund_choices template from master.")
    ap.add_argument("--master", required=True, help="latestNAV_Reports.xlsx")
    ap.add_argument(
        "--out",
        default=os.path.join("data", "dummy_clients", "fund_choices_template.xlsx"),
        help="Output path",
    )
    ns = ap.parse_args()
    p = build_fund_choices_template(ns.master, ns.out)
    print("Wrote", os.path.abspath(ns.out), p)
