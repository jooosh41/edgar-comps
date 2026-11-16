"""Run the extractor across many sectors and report coverage. Usage: python scripts/stress_test.py"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from edgar_comps import facts
from edgar_comps.loader import load_company

SECTORS = {
    "Tech": ["AAPL", "MSFT", "GOOGL", "NVDA", "META", "ORCL"],
    "Consumer": ["WMT", "COST", "KO", "PEP", "MCD", "NKE", "SBUX", "HD"],
    "Internet": ["AMZN", "UBER", "ABNB", "TSLA"],
    "Industrial": ["BA", "CAT", "F", "DAL", "UPS"],
    "Health": ["JNJ", "PFE", "UNH"],
    "Energy/Utility": ["XOM", "CVX", "NEE"],
    "Telecom": ["T", "VZ"],
    "Financial": ["JPM", "GS", "V"],
    "Real estate": ["O"],
    "Edge cases": ["BRK-B", "SPOT", "TSM"],
}


def run():
    rows = []
    for sector, tickers in SECTORS.items():
        for t in tickers:
            try:
                c = load_company(t)
            except Exception as error:                      # report, never crash the sweep
                rows.append((sector, t, "FAILED", str(error)[:70], None))
                continue
            gaps = facts.missing_items(c)
            notes = facts.fallbacks_used(c)
            status = "gaps" if gaps else ("approx" if notes else "clean")
            rows.append((sector, t, status, "; ".join(gaps) or "; ".join(n.split(" (")[0] for n in notes), c))
    return rows


if __name__ == "__main__":
    rows = run()
    for sector, t, status, detail, c in rows:
        extra = ""
        if c:
            rev, ebit = c.latest("revenue"), c.latest("operating_income")
            extra = f"FY {c.year_ends[-1]}  rev {rev / 1e6:>10,.0f}  EBIT {'n/a' if ebit is None else format(ebit / 1e6, ',.0f'):>9}  yrs {len(c.year_ends)}"
        print(f"{sector:15s} {t:6s} {status:7s} {extra}  {detail}")
    n = len(rows)
    count = lambda status: sum(r[2] == status for r in rows)
    print(f"\n{n} tested: {count('clean')} clean, {count('approx')} usable with a disclosed approximation, {count('gaps')} with gaps, {count('FAILED')} failed")
