"""
config.py — All paths, thresholds and constants in one place.
Edit this file to configure the system without touching any other module.
"""

import os

# ── File Paths ─────────────────────────────────────────────────────────────────
ISIN_MASTER_FILE     = os.path.join("data", "latestNAV_Reports.xlsx")
AMC_DISCLOSURE_DIR   = os.path.join("data", "amc_disclosures")
CLIENT_DIR           = os.path.join("data", "client_holdings")
OUTPUT_CSV           = os.path.join("output", "transactions_bulk.csv")

# Morningstar ID cache — maps ISIN → Morningstar fund ID (e.g. "f00001q819")
# Populated by morningstar.py and persisted here to avoid re-scraping.
MORNINGSTAR_ID_CACHE = os.path.join("output", "morningstar_id_cache.json")

# ── Transaction Rules ──────────────────────────────────────────────────────────
MIN_TRANSACTION_AMT = 5_000    # Suppress transactions below this amount (₹)
TOP_N_FUNDS         = 3        # Max funds suggested per direction (buy / sell)

# ── Rebalancing Archetypes ─────────────────────────────────────────────────────
# equity + defensive + other == 1.0
ARCHETYPES = {
    "Averse":     {"equity": 0.35, "defensive": 0.60, "other": 0.05},
    "Moderate":   {"equity": 0.50, "defensive": 0.45, "other": 0.05},
    "Aggressive": {"equity": 0.65, "defensive": 0.30, "other": 0.05},
}

# How "defensive %" is computed in portfolio / vs targets:
#   "debt_cash"      — debt + cash only (other is separate bucket)
#   "debt_cash_other" — debt + cash + other
DEFENSIVE_BUCKETS = "debt_cash"

OTHERS_MIN        = 0.05   # minimum others allocation (portfolio weight)
OTHERS_MAX_RATIO  = 0.50   # max others weight = defensive_target * this ratio

# ── Age-based glide path (bulk_run + MFD pack when enabled) ───────────────────
# Equity % (percentage points) = 100 - k * age, then converted to a fraction and
# clamped. Other bucket target is flat TARGET_OTHER_DEFAULT; defensive = remainder.
GLIDE_PATH_K = {
    "Moderate": 1.0,
    "Aggressive": 0.5,
    "Averse": 1.5,
}
GLIDE_EQUITY_MIN = 0.20
GLIDE_EQUITY_MAX = 0.90
TARGET_OTHER_DEFAULT = 0.05  # flat other target with glide path (same as typical archetype)

# Optional default path for dummy / tests (override via CLI or INI)
CLIENT_AGES_FILE = os.path.join("data", "dummy_clients", "client_ages.xlsx")

# New-fund sleeves (Final cols U–X) — fund_choices.py validation
NEW_SCHEME_MIN_EQUITY_FRAC = 0.85
NEW_SCHEME_MIN_DEFENSIVE_FRAC = 0.80  # final_debt + final_cash
NEW_SCHEME_MIN_OTHER_FRAC = 0.80

# Optional path for bulk_run / MFD pack when allow_new_funds=true (FundChoices sheet)
FUND_CHOICES_FILE = os.path.join("data", "dummy_clients", "fund_choices.xlsx")

# ── AMC Parser: Fuzzy Match Settings ──────────────────────────────────────────
# STRICT threshold for AMC file → ISIN master matching.
# Keeps only matches where BOTH:
#   1. Token overlap score ≥ FUZZY_MIN_SCORE
#   2. AMC name in scheme name starts with known AMC prefix
# This prevents cross-AMC contamination (e.g. Motilal "Nifty 500 Momentum 50"
# matching Groww "Nifty 500 Momentum 50 ETF" — same index, different AMC).
FUZZY_MIN_SCORE    = 0.55     # Minimum Jaccard token overlap score
REQUIRE_AMC_MATCH  = True     # If True, AMC name in scheme must match AMC file source

# Noise words stripped before token overlap scoring
FUZZY_NOISE_WORDS  = {
    "fund", "plan", "direct", "regular", "growth", "idcw", "the", "an", "a",
    "of", "and", "in", "for", "formerly", "known", "as", "scheme", "option",
    "dividend", "payout", "reinvestment", "bonus", "quarterly", "monthly",
    "annual", "weekly", "half", "yearly", "open", "ended", "close", "etf",
}

# ── Hybrid Fund Categories (use AMC actual data when available) ───────────────
HYBRID_CATEGORIES = {
    "Open Ended Scheme (Hybrid Scheme - Aggressive Hybrid Fund)",
    "Open Ended Scheme (Hybrid Scheme - Arbitrage Fund)",
    "Open Ended Scheme (Hybrid Scheme - Balanced Hybrid Fund)",
    "Open Ended Scheme (Hybrid Scheme - Conservative Hybrid Fund)",
    "Open Ended Scheme (Hybrid Scheme - Dynamic Asset Allocation or Balanced Advantage)",
    "Open Ended Scheme (Hybrid Scheme - Equity Savings)",
    "Open Ended Scheme (Hybrid Scheme - Multi Asset Allocation)",
    "Open Ended Scheme (Other Scheme - FoF Domestic)",
    "Open Ended Scheme (Other Scheme - FoF Overseas)",
    "Open Ended Scheme (Solution Oriented Scheme - Children's Fund)",
    "Open Ended Scheme (Solution Oriented Scheme - Retirement Fund)",
}

# ── Morningstar Scraper Settings ───────────────────────────────────────────────
MORNINGSTAR_BASE      = "https://www.morningstar.in/mutualfunds"
MORNINGSTAR_SEARCH    = "https://www.morningstar.in/handlers/autocomplete.ashx"
MORNINGSTAR_RATE_LIMIT_SEC = 1.5   # Pause between requests (be a polite scraper)
MORNINGSTAR_MAX_RETRIES    = 2
