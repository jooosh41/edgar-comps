"""Command line: python -m edgar_comps NFLX DIS CMCSA --target NFLX"""
import argparse
import csv
import sys
from pathlib import Path

from . import facts, market, workbook
from .loader import load_company

ROOT = Path(__file__).resolve().parent.parent


def read_prices(path):
    """CSV with header ticker,price,date,source -> {ticker: (price, date, source)}."""
    prices = {}
    with open(path, newline="") as handle:
        for row in csv.DictReader(handle):
            prices[row["ticker"].strip().upper()] = (float(row["price"]), row.get("date", ""), row.get("source", ""))
    return prices


def read_assumptions(path):
    """CSV with header key,value,note -> {key: (value, note)}. Keys are listed in workbook.ASSUMPTION_ROWS."""
    overrides = {}
    with open(path, newline="") as handle:
        for row in csv.DictReader(handle):
            key = row["key"].strip()
            if key not in workbook.ASSUMPTION_ROWS:
                sys.exit(f"{path}: unknown assumption '{key}'. Valid keys: {', '.join(workbook.ASSUMPTION_ROWS)}")
            overrides[key] = (float(row["value"]), row.get("note", "").strip())
    return overrides


def main(argv=None):
    parser = argparse.ArgumentParser(prog="edgar_comps", description=__doc__)
    parser.add_argument("tickers", nargs="*", help="peer tickers; default is the --peers file")
    parser.add_argument("--peers", default="imax", help="name of a file in peers/, without .txt (imax or media)")
    parser.add_argument("--target", help="company to run the DCF on; default is the first ticker")
    parser.add_argument("--out", default=str(ROOT / "output" / "comps.xlsx"))
    parser.add_argument("--prices", help="optional CSV: ticker,price,date,source")
    parser.add_argument("--market", action="store_true", help="fetch share prices, beta and the risk-free rate")
    parser.add_argument("--assumptions", help="CSV of DCF overrides: key,value,note (default: assumptions/<TARGET>.csv if it exists)")
    parser.add_argument("--years", type=int, default=5, choices=range(1, 6))
    args = parser.parse_args(argv)

    tickers = args.tickers or [
        line.split("#")[0].strip() for line in (ROOT / "peers" / f"{args.peers}.txt").read_text().splitlines()
        if line.split("#")[0].strip()]
    tickers = [t.upper() for t in tickers]
    target = (args.target or tickers[0]).upper()
    if target not in tickers:
        tickers.insert(0, target)

    companies = []
    for ticker in tickers:
        try:
            company = load_company(ticker, years=args.years)
        except (KeyError, ValueError) as error:
            print(f"  skipped {ticker}: {error}", file=sys.stderr)
            continue
        gaps = facts.missing_items(company)
        print(f"  {ticker:6s} {len(company.year_ends)} years, latest {company.year_ends[-1]}"
              + (f"  MISSING: {'; '.join(gaps)}" if gaps else "")
              + (f"  APPROX: {len(facts.fallbacks_used(company))}" if facts.fallbacks_used(company) else ""))
        companies.append(company)

    by_ticker = {c.ticker: c for c in companies}
    if target not in by_ticker:
        sys.exit(f"Target {target} could not be loaded, so there is nothing to value.")
    core_gaps = [g for g in facts.missing_items(by_ticker[target]) if not g.startswith("Debt")]
    if core_gaps:
        sys.exit(f"Target {target} is missing {', '.join(core_gaps)}; a DCF on it would be wrong. Pick another target.")

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    market_data = None
    if args.market:
        market_data = market.gather([c.ticker for c in companies], target)
        for problem in market_data["errors"]:
            print(f"  market data: {problem}", file=sys.stderr)
        print(f"  market data: {len(market_data['prices'])} prices, beta {'ok' if market_data['beta'] else 'missing'}, "
              f"risk-free {'ok' if market_data['risk_free'] else 'missing'}")
    assumptions_path = Path(args.assumptions) if args.assumptions else ROOT / "assumptions" / f"{target}.csv"
    assumptions = read_assumptions(assumptions_path) if assumptions_path.exists() else None
    if assumptions:
        print(f"  assumptions: {len(assumptions)} overrides from {assumptions_path.name}")
    workbook.build(companies, target, out, read_prices(args.prices) if args.prices else None, market_data, assumptions)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
