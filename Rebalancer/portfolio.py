"""
portfolio.py — Portfolio calculation and transaction generation.

DESIGN PRINCIPLE: Every intermediate result is inspectable in Excel.
Final equity/debt/cash/other per ISIN come from cols U-X of latestNAV_Reports.xlsx
(written by add_final_columns.py). Col T is MS Total % (update_morningstar.py).

BUCKET DEFINITIONS:
  Final Equity %  (col U) — equities, REITs, foreign equities
  Final Debt %    (col V) — bonds, G-Secs, money market
  Final Cash %    (col W) — cash/MM
  Final Other %   (col X) — gold, silver, commodity, unclassified

DEFENSIVE_BUCKETS (config):
  "debt_cash"       — defensive % = debt + cash (other separate)
  "debt_cash_other" — defensive % = debt + cash + other

REBALANCING:
  1) Others band: [OTHERS_MIN, target_def * OTHERS_MAX_RATIO]
  2) Equity vs defensive (per DEFENSIVE_BUCKETS) toward archetype targets.

  With new_cash == 0, recycling legs are built so total BUY INR == total SELL INR
  (paired switches). net_flow_inr in the summary should be 0. With new_cash > 0,
  net_flow_inr equals deployed fresh cash (new inflow); see summary fields.
"""

import pandas as pd
from config import (
    ARCHETYPES,
    DEFENSIVE_BUCKETS,
    GLIDE_EQUITY_MAX,
    GLIDE_EQUITY_MIN,
    GLIDE_PATH_K,
    MIN_TRANSACTION_AMT,
    OTHERS_MAX_RATIO,
    OTHERS_MIN,
    TARGET_OTHER_DEFAULT,
    TOP_N_FUNDS,
)

COL_FINAL_EQUITY = "final_equity"
COL_FINAL_DEBT   = "final_debt"
COL_FINAL_CASH   = "final_cash"
COL_FINAL_OTHER  = "final_other"


def load_master_for_portfolio(master_path: str) -> pd.DataFrame:
    """
    Load latestNAV_Reports.xlsx with only the columns needed for portfolio
    computation. Expects Final columns (U-X) from add_final_columns.py.
    """
    df = pd.read_excel(master_path, sheet_name="latestNAV_Reports")
    raw_cols = list(df.columns)

    col_map = {
        "isin":         raw_cols[4],   # E
        "scheme_name":  raw_cols[3],   # D
        "nav":          raw_cols[6],   # G
        "final_equity": raw_cols[20],  # U
        "final_debt":   raw_cols[21],  # V
        "final_cash":   raw_cols[22],  # W
        "final_other":  raw_cols[23],  # X
    }

    result = pd.DataFrame()
    for new_name, old_name in col_map.items():
        result[new_name] = df[old_name]

    result["isin"] = result["isin"].astype(str).str.strip()
    result = result[
        result["isin"].notna()
        & (result["isin"] != "nan")
        & (result["isin"].str.len() >= 10)
    ].drop_duplicates("isin").reset_index(drop=True)

    return result


def compute_glide_target_mix(
    age: float,
    risk_preference: str,
    *,
    others_target: float | None = None,
) -> dict[str, float]:
    """
    Continuous glide: equity %% (percentage points) = 100 - k * age, k from risk preference:
      Moderate=1, Aggressive=0.5, Averse=1.5.

    Returns dict with keys equity, defensive, other (fractions summing to 1).
    ``other`` is fixed at TARGET_OTHER_DEFAULT (or ``others_target`` override).
    ``defensive`` = 1 - equity - other after equity is clamped and capped by room.
    """
    pref = risk_preference.strip().title()
    if pref not in GLIDE_PATH_K:
        raise KeyError(
            f"Unknown risk_preference {risk_preference!r}; "
            f"expected one of {sorted(GLIDE_PATH_K)}"
        )
    k = GLIDE_PATH_K[pref]
    other = float(TARGET_OTHER_DEFAULT if others_target is None else others_target)
    other = max(0.0, min(other, 0.95))

    eq = (100.0 - k * float(age)) / 100.0
    eq = max(GLIDE_EQUITY_MIN, min(GLIDE_EQUITY_MAX, eq))

    eq_cap = 1.0 - other
    if eq > eq_cap:
        eq = eq_cap
    defensive = 1.0 - eq - other
    if defensive < -1e-9:
        eq = max(GLIDE_EQUITY_MIN, min(GLIDE_EQUITY_MAX, eq_cap))
        defensive = 1.0 - eq - other
    defensive = max(0.0, defensive)

    s = eq + defensive + other
    if s > 0 and abs(s - 1.0) > 1e-9:
        eq /= s
        defensive /= s
        other /= s

    return {
        "equity": round(eq, 6),
        "defensive": round(defensive, 6),
        "other": round(other, 6),
    }


def _is_equity(row: pd.Series) -> bool:
    v = row.get(COL_FINAL_EQUITY)
    return pd.notna(v) and float(v) >= 0.5


def _is_other(row: pd.Series) -> bool:
    v = row.get(COL_FINAL_OTHER)
    return pd.notna(v) and float(v) >= 0.3


def _is_defensive(row: pd.Series) -> bool:
    return (not _is_equity(row)) and (not _is_other(row))


def _top_n(holdings: pd.DataFrame, pred, n: int = TOP_N_FUNDS) -> pd.DataFrame:
    if holdings.empty:
        return holdings
    mask = holdings.apply(lambda r: pred(r), axis=1)
    sub = holdings[mask]
    sub = sub[sub["value"] > 0].sort_values("value", ascending=False)
    return sub.head(n)


def compute_portfolio(holdings, master_df: pd.DataFrame) -> dict:
    """
    holdings: list of (isin, units) or DataFrame [isin, units]
    master_df: output of load_master_for_portfolio()
    """
    if isinstance(holdings, (list, tuple)):
        hdf = pd.DataFrame(holdings, columns=["isin", "units"])
    else:
        hdf = holdings.copy()
        hdf.columns = [c.lower().strip() for c in hdf.columns]

    hdf["isin"] = hdf["isin"].astype(str).str.strip()
    hdf["units"] = pd.to_numeric(hdf["units"], errors="coerce").fillna(0)

    merged = hdf.merge(master_df, on="isin", how="left")
    merged["nav"] = pd.to_numeric(merged["nav"], errors="coerce").fillna(0)
    merged["scheme_name"] = merged["scheme_name"].fillna(merged["isin"])

    for col in [COL_FINAL_EQUITY, COL_FINAL_DEBT, COL_FINAL_CASH, COL_FINAL_OTHER]:
        merged[col] = pd.to_numeric(merged[col], errors="coerce")

    missing_mask = merged[COL_FINAL_EQUITY].isna()
    merged.loc[missing_mask, COL_FINAL_EQUITY] = 0.5
    merged.loc[missing_mask, COL_FINAL_DEBT] = 0.5
    merged.loc[missing_mask, COL_FINAL_CASH] = 0.0
    merged.loc[missing_mask, COL_FINAL_OTHER] = 0.0

    merged["value"] = merged["units"] * merged["nav"]
    merged["equity_value"] = merged["value"] * merged[COL_FINAL_EQUITY]
    merged["debt_value"] = merged["value"] * merged[COL_FINAL_DEBT]
    merged["cash_value"] = merged["value"] * merged[COL_FINAL_CASH]
    merged["other_value"] = merged["value"] * merged[COL_FINAL_OTHER]

    total = merged["value"].sum()
    eq_val = merged["equity_value"].sum()
    dbt_val = merged["debt_value"].sum()
    cash_val = merged["cash_value"].sum()
    oth_val = merged["other_value"].sum()

    if DEFENSIVE_BUCKETS == "debt_cash":
        def_val = dbt_val + cash_val
    else:
        def_val = dbt_val + cash_val + oth_val

    def pct(v):
        return round(v / total, 4) if total > 0 else 0.0

    return {
        "total_value": round(total, 2),
        "equity_value": round(eq_val, 2),
        "debt_value": round(dbt_val, 2),
        "cash_value": round(cash_val, 2),
        "other_value": round(oth_val, 2),
        "equity_pct": pct(eq_val),
        "debt_pct": pct(dbt_val),
        "cash_pct": pct(cash_val),
        "other_pct": pct(oth_val),
        "defensive_pct": pct(def_val),
        "holdings": merged,
    }


def _current_defensive_value(portfolio: dict) -> float:
    if DEFENSIVE_BUCKETS == "debt_cash":
        return portfolio["debt_value"] + portfolio["cash_value"]
    return portfolio["debt_value"] + portfolio["cash_value"] + portfolio["other_value"]


def _pro_rata_line_amounts(
    pool: pd.DataFrame,
    total_amt: float,
    min_amt: float,
) -> list[tuple[pd.Series, float]]:
    """
    Split total_amt across pool rows by holding value weights.
    Returned amounts sum to total_amt (2 d.p.), each 0 or >= min_amt.
    """
    if pool.empty or total_amt <= 0:
        return []
    total_amt = round(float(total_amt), 2)
    if total_amt < min_amt:
        return []
    pool_total = float(pool["value"].sum())
    n = len(pool)
    rows_weights: list[tuple[pd.Series, float]] = []
    for _, row in pool.iterrows():
        w = float(row["value"]) / pool_total if pool_total > 0 else 1.0 / n
        rows_weights.append((row, w))
    amts = [round(total_amt * w, 2) for _, w in rows_weights]
    amts = [x if x >= min_amt else 0.0 for x in amts]
    s = round(sum(amts), 2)
    gap = round(total_amt - s, 2)
    if abs(gap) >= 0.005:
        idx = max(range(len(amts)), key=lambda i: amts[i])
        if amts[idx] <= 0:
            idx = max(range(len(amts)), key=lambda i: rows_weights[i][1])
        amts[idx] = round(amts[idx] + gap, 2)
    out = [(rows_weights[i][0], amts[i]) for i in range(len(amts)) if amts[i] >= min_amt]
    if not out and total_amt >= min_amt:
        best_i = max(range(len(rows_weights)), key=lambda i: rows_weights[i][1])
        out = [(rows_weights[best_i][0], total_amt)]
    return out


def _surplus_values(
    rv_eq: float,
    rv_def: float,
    rv_oth: float,
    target_eq_val: float,
    target_def_val: float,
    target_oth_val: float,
) -> tuple[float, float, float]:
    """INR surplus in each bucket vs archetype / glide targets (not below zero)."""
    return (
        max(0.0, rv_eq - target_eq_val),
        max(0.0, rv_def - target_def_val),
        max(0.0, rv_oth - target_oth_val),
    )


def _pick_others_band_funding_sell(
    holdings: pd.DataFrame,
    surplus_eq: float,
    surplus_def: float,
    surplus_oth: float,
) -> tuple[pd.DataFrame, str, float]:
    """
    For raising the *others* allocation, sell only from a bucket in surplus.
    Preference: equity → defensive → other (typical: trim equity before debt).
    Returns (pool, \"equity\"|\"defensive\"|\"other\"|\"\", cap_inr).
    """
    if surplus_eq >= MIN_TRANSACTION_AMT:
        po = _top_n(holdings, _is_equity)
        if not po.empty:
            return po, "equity", surplus_eq
    if surplus_def >= MIN_TRANSACTION_AMT:
        po = _top_n(holdings, _is_defensive)
        if not po.empty:
            return po, "defensive", surplus_def
    if surplus_oth >= MIN_TRANSACTION_AMT:
        po = _top_n(holdings, _is_other)
        if not po.empty:
            return po, "other", surplus_oth
    return pd.DataFrame(), "", 0.0


def _append_weighted(
    transactions: list[dict],
    pairs: list[tuple[pd.Series, float]],
    action_label: str,
) -> float:
    total = 0.0
    for row, amt in pairs:
        transactions.append({
            "isin": row["isin"],
            "scheme_name": row.get("scheme_name", ""),
            "action": action_label,
            "amount_inr": amt,
        })
        total += amt
    return round(total, 2)


def generate_transactions(
    portfolio: dict,
    archetype: str,
    new_cash: float = 0,
    new_scheme_pools: dict[str, list[tuple[str, float]]] | None = None,
    master_df: pd.DataFrame | None = None,
    target_mix: dict | None = None,
) -> tuple[list[dict], dict]:
    """
    If ``target_mix`` is set (keys equity, defensive, other — fractions), it
    overrides ``ARCHETYPES[archetype]`` for targets. The ``archetype`` argument
    (e.g. \"Glide:Moderate\") is kept in ``summary[\"archetype\"]`` for logging;
    CSV/Excel export uses ``bulk_run.display_risk_type`` for column **target_policy**.

    ``new_scheme_pools`` keys equity / defensive / other map to
    ``[(isin, weight), ...]`` for weighted BUY pools (see fund_choices.py).
    """
    if target_mix is not None:
        target_eq = float(target_mix["equity"])
        target_def = float(target_mix["defensive"])
        target_other = float(target_mix["other"])
    else:
        target_eq = ARCHETYPES[archetype]["equity"]
        target_def = ARCHETYPES[archetype]["defensive"]
        target_other = ARCHETYPES[archetype]["other"]
    others_max = target_def * OTHERS_MAX_RATIO

    holdings = portfolio["holdings"].copy()
    total_val = portfolio["total_value"]
    new_total = total_val + new_cash

    target_eq_val = new_total * target_eq
    target_def_val = new_total * target_def
    target_oth_val = new_total * target_other

    rv_eq = float(portfolio["equity_value"])
    rv_def = float(_current_defensive_value(portfolio))
    rv_oth = float(portfolio["other_value"])

    transactions: list[dict] = []

    def distribute_exact(pool: pd.DataFrame, amount: float, action_label: str) -> float:
        pairs = _pro_rata_line_amounts(pool, amount, MIN_TRANSACTION_AMT)
        return _append_weighted(transactions, pairs, action_label)

    def distribute_pair(
        sell_pool: pd.DataFrame,
        buy_pool: pd.DataFrame,
        amount: float,
        sell_label: str,
        buy_label: str,
    ) -> float:
        """Recycle cash: total SELL INR == total BUY INR (2 d.p.). Returns flow or 0."""
        if sell_pool.empty or buy_pool.empty or amount <= 0:
            return 0.0
        sell_pairs = _pro_rata_line_amounts(sell_pool, amount, MIN_TRANSACTION_AMT)
        if not sell_pairs:
            return 0.0
        sell_sum = round(sum(a for _, a in sell_pairs), 2)
        _append_weighted(transactions, sell_pairs, sell_label)
        buy_pairs = _pro_rata_line_amounts(buy_pool, sell_sum, MIN_TRANSACTION_AMT)
        if not buy_pairs:
            b0 = buy_pool.iloc[0]
            one_row = pd.DataFrame([b0])
            buy_pairs = _pro_rata_line_amounts(one_row, sell_sum, MIN_TRANSACTION_AMT)
        if not buy_pairs:
            return sell_sum
        buy_sum = round(sum(a for _, a in buy_pairs), 2)
        gap = round(sell_sum - buy_sum, 2)
        if abs(gap) >= 0.005 and buy_pairs:
            row0, a0 = buy_pairs[-1]
            buy_pairs[-1] = (row0, round(a0 + gap, 2))
        _append_weighted(transactions, buy_pairs, buy_label)
        return sell_sum

    def buy_pool_equity():
        if new_scheme_pools and master_df is not None:
            entries = new_scheme_pools.get("equity") or []
            if entries:
                po = _new_scheme_pool_from_entries(master_df, entries, "equity")
                if not po.empty:
                    return po
        return _top_n(holdings, _is_equity)

    def buy_pool_defensive():
        if new_scheme_pools and master_df is not None:
            entries = new_scheme_pools.get("defensive") or []
            if entries:
                po = _new_scheme_pool_from_entries(master_df, entries, "defensive")
                if not po.empty:
                    return po
        return _top_n(holdings, _is_defensive)

    def buy_pool_other():
        if new_scheme_pools and master_df is not None:
            entries = new_scheme_pools.get("other") or []
            if entries:
                po = _new_scheme_pool_from_entries(master_df, entries, "other")
                if not po.empty:
                    return po
        return _top_n(holdings, _is_other)

    def sell_pool_equity():
        return _top_n(holdings, _is_equity)

    def sell_pool_defensive():
        return _top_n(holdings, _is_defensive)

    # Cash from new_cash actually placed into "other" (may be 0 if no other pool).
    deployed_new_cash_to_others = 0.0

    # ── Step 1: Others band ─────────────────────────────────────────────────
    cur_other_pct = portfolio["other_pct"]

    if cur_other_pct < OTHERS_MIN - 1e-9:
        need_inr = (OTHERS_MIN - cur_other_pct) * new_total
        if need_inr >= MIN_TRANSACTION_AMT:
            planned_cash_to_others = min(new_cash, need_inr) if new_cash > 0 else 0.0
            if planned_cash_to_others >= MIN_TRANSACTION_AMT:
                t = distribute_exact(
                    buy_pool_other(),
                    planned_cash_to_others,
                    "BUY (others band → add other)",
                )
                rv_oth += t
                deployed_new_cash_to_others = t
            residual = need_inr - deployed_new_cash_to_others
            if residual >= MIN_TRANSACTION_AMT:
                seq, sdf, sot = _surplus_values(
                    rv_eq, rv_def, rv_oth,
                    target_eq_val, target_def_val, target_oth_val,
                )
                sell_pool, sold_from, cap_inr = _pick_others_band_funding_sell(
                    holdings, seq, sdf, sot,
                )
                eff = min(residual, cap_inr)
                if sell_pool.empty or eff < MIN_TRANSACTION_AMT:
                    eff = 0.0
                else:
                    pool_cap = float(sell_pool["value"].sum())
                    eff = min(eff, pool_cap)
                if eff >= MIN_TRANSACTION_AMT:
                    flow = distribute_pair(
                        sell_pool,
                        buy_pool_other(),
                        eff,
                        "SELL (→ fund others band)",
                        "BUY (others band → add other)",
                    )
                    if sold_from == "defensive":
                        rv_def -= flow
                    elif sold_from == "equity":
                        rv_eq -= flow
                    else:
                        rv_oth -= flow
                    rv_oth += flow

    elif cur_other_pct > others_max + 1e-9:
        excess_inr = (cur_other_pct - others_max) * new_total
        if excess_inr >= MIN_TRANSACTION_AMT:
            sell_pool_o = _top_n(holdings, _is_other)
            sell_pairs = _pro_rata_line_amounts(
                sell_pool_o, excess_inr, MIN_TRANSACTION_AMT
            )
            st = round(sum(a for _, a in sell_pairs), 2) if sell_pairs else 0.0
            if st >= MIN_TRANSACTION_AMT:
                _append_weighted(
                    transactions, sell_pairs, "SELL (others band → reduce other)"
                )
                need_eq = max(0.0, target_eq_val - rv_eq)
                need_def = max(0.0, target_def_val - rv_def)
                tot_need = need_eq + need_def
                if tot_need >= MIN_TRANSACTION_AMT:
                    buy_eq_amt = round(st * (need_eq / tot_need), 2)
                    buy_def_amt = round(st - buy_eq_amt, 2)
                else:
                    den = target_eq + target_def
                    if den <= 0:
                        den = 1.0
                    buy_eq_amt = round(st * (target_eq / den), 2)
                    buy_def_amt = round(st - buy_eq_amt, 2)
                if buy_def_amt > 0 and buy_def_amt < MIN_TRANSACTION_AMT:
                    buy_eq_amt = st
                    buy_def_amt = 0.0
                elif buy_eq_amt > 0 and buy_eq_amt < MIN_TRANSACTION_AMT:
                    buy_def_amt = st
                    buy_eq_amt = 0.0
                te = distribute_exact(
                    buy_pool_equity(), buy_eq_amt, "BUY (others band → add equity)"
                )
                td = distribute_exact(
                    buy_pool_defensive(), buy_def_amt, "BUY (others band → add defensive)"
                )
                slip = round(st - te - td, 2)
                if slip >= MIN_TRANSACTION_AMT:
                    if buy_eq_amt >= buy_def_amt:
                        te += distribute_exact(
                            buy_pool_equity(), slip, "BUY (others band → add equity)"
                        )
                    else:
                        td += distribute_exact(
                            buy_pool_defensive(), slip, "BUY (others band → add defensive)"
                        )
                rv_oth -= st
                rv_eq += te
                rv_def += td

    remaining_new_cash = max(0.0, new_cash - deployed_new_cash_to_others)

    # ── Step 2: New cash → equity vs defensive ─────────────────────────────
    cash_to_equity = 0.0
    cash_to_def = 0.0
    if remaining_new_cash > 0:
        eq_gap = target_eq_val - rv_eq
        def_gap = target_def_val - rv_def
        if eq_gap > 0:
            cash_to_equity = min(remaining_new_cash, eq_gap)
            cash_to_def = remaining_new_cash - cash_to_equity
        else:
            cash_to_def = min(remaining_new_cash, max(0.0, def_gap))
            cash_to_equity = remaining_new_cash - cash_to_def

    if cash_to_equity > MIN_TRANSACTION_AMT:
        rv_eq += distribute_exact(
            buy_pool_equity(), cash_to_equity, "BUY (new cash → equity)"
        )
    if cash_to_def > MIN_TRANSACTION_AMT:
        rv_def += distribute_exact(
            buy_pool_defensive(), cash_to_def, "BUY (new cash → defensive)"
        )

    residual_eq_gap = target_eq_val - rv_eq
    switch_amount = abs(residual_eq_gap)

    # ── Step 3: Switch equity ↔ defensive (sell only surplus buckets) ─────
    seq3, sdf3, sot3 = _surplus_values(
        rv_eq, rv_def, rv_oth,
        target_eq_val, target_def_val, target_oth_val,
    )
    if switch_amount > MIN_TRANSACTION_AMT:
        if residual_eq_gap > 0:
            spool = pd.DataFrame()
            sold_from = ""
            cap_sw = 0.0
            if sdf3 >= MIN_TRANSACTION_AMT:
                spool = sell_pool_defensive()
                cap_sw = min(switch_amount, sdf3)
                sold_from = "defensive"
            elif sot3 >= MIN_TRANSACTION_AMT:
                spool = _top_n(holdings, _is_other)
                cap_sw = min(switch_amount, sot3)
                sold_from = "other"
            if sold_from and not spool.empty:
                pool_max = float(spool["value"].sum())
                eff_sw = min(cap_sw, pool_max)
                if eff_sw >= MIN_TRANSACTION_AMT:
                    flow = distribute_pair(
                        spool,
                        buy_pool_equity(),
                        eff_sw,
                        "SELL (rebalance → reduce defensive)"
                        if sold_from == "defensive"
                        else "SELL (rebalance → reduce other)",
                        "BUY  (rebalance → add equity)",
                    )
                    if sold_from == "defensive":
                        rv_def -= flow
                    else:
                        rv_oth -= flow
                    rv_eq += flow
        else:
            spool = pd.DataFrame()
            sold_from = ""
            cap_sw = 0.0
            if seq3 >= MIN_TRANSACTION_AMT:
                spool = sell_pool_equity()
                cap_sw = min(switch_amount, seq3)
                sold_from = "equity"
            elif sot3 >= MIN_TRANSACTION_AMT:
                spool = _top_n(holdings, _is_other)
                cap_sw = min(switch_amount, sot3)
                sold_from = "other"
            if sold_from and not spool.empty:
                pool_max = float(spool["value"].sum())
                eff_sw = min(cap_sw, pool_max)
                if eff_sw >= MIN_TRANSACTION_AMT:
                    flow = distribute_pair(
                        spool,
                        buy_pool_defensive(),
                        eff_sw,
                        "SELL (rebalance → reduce equity)"
                        if sold_from == "equity"
                        else "SELL (rebalance → reduce other)",
                        "BUY  (rebalance → add defensive)",
                    )
                    if sold_from == "equity":
                        rv_eq -= flow
                    else:
                        rv_oth -= flow
                    rv_def += flow

    total_buy = round(
        sum(
            t["amount_inr"]
            for t in transactions
            if str(t.get("action", "")).startswith("BUY")
        ),
        2,
    )
    total_sell = round(
        sum(
            t["amount_inr"]
            for t in transactions
            if str(t.get("action", "")).startswith("SELL")
        ),
        2,
    )
    summary = {
        "archetype": archetype,
        "target_equity_pct": target_eq,
        "target_defensive_pct": target_def,
        "target_other_pct": target_other,
        "current_equity_pct": portfolio["equity_pct"],
        "current_defensive_pct": portfolio["defensive_pct"],
        "current_other_pct": portfolio["other_pct"],
        "others_min": OTHERS_MIN,
        "others_max": round(others_max, 4),
        "total_portfolio": total_val,
        "new_cash": new_cash,
        "deployed_new_cash_to_others_inr": round(deployed_new_cash_to_others, 2),
        "switch_amount": round(switch_amount, 2),
        "cash_to_equity": round(cash_to_equity, 2),
        "cash_to_defensive": round(cash_to_def, 2),
        "total_buy_inr": total_buy,
        "total_sell_inr": total_sell,
        "net_flow_inr": round(total_buy - total_sell, 2),
    }
    return transactions, summary


def _new_scheme_pool_from_entries(
    master_df: pd.DataFrame,
    entries: list[tuple[str, float]],
    side: str,
) -> pd.DataFrame:
    """
    Build a synthetic pool in ``entries`` order with ``value`` = Excel weight.
    Drops rows that fail ``_is_equity`` / ``_is_defensive`` / ``_is_other``.
    """
    parts: list[pd.DataFrame] = []
    seen: set[str] = set()
    for isin, wt in entries:
        if wt <= 0:
            continue
        isin = str(isin).strip()
        if not isin or isin in seen:
            continue
        m = master_df[master_df["isin"] == isin]
        if m.empty:
            continue
        part = m.iloc[[0]].copy()
        part["value"] = float(wt)
        parts.append(part)
        seen.add(isin)
    if not parts:
        return pd.DataFrame()
    sub = pd.concat(parts, ignore_index=True)
    if side == "equity":
        mask = sub.apply(_is_equity, axis=1)
    elif side == "other":
        mask = sub.apply(_is_other, axis=1)
    else:
        mask = sub.apply(_is_defensive, axis=1)
    sub = sub[mask]
    return sub.head(TOP_N_FUNDS)
