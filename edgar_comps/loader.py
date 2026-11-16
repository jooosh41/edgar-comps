"""Load one company by ticker, falling back to a predecessor entity when needed.

A corporate reorganisation can move a ticker to a new SEC registrant (new CIK) that has
no 10-K history yet. peers/predecessors.txt maps such tickers to the old CIK.
"""
from pathlib import Path

from . import edgar, facts

ROOT = Path(__file__).resolve().parent.parent


def _predecessors():
    table = {}
    path = ROOT / "peers" / "predecessors.txt"
    if path.exists():
        for line in path.read_text().splitlines():
            parts = line.split("#")[0].split()
            if len(parts) == 2:
                table[parts[0].upper()] = int(parts[1])
    return table


def load_company(ticker, years=5):
    """Return a facts.Company. Raises KeyError (unknown ticker) or ValueError (no annual data)."""
    ticker = ticker.upper()
    cik, name = edgar.lookup(ticker)
    try:
        return facts.build_company(ticker, name, cik, edgar.company_facts(cik), years=years)
    except ValueError:
        old_cik = _predecessors().get(ticker)
        if old_cik is None:
            raise
        raw = edgar.company_facts(old_cik)
        company = facts.build_company(ticker, name, old_cik, raw, years=years)
        company.notes.append(f"History from predecessor registrant {raw.get('entityName', '')} (CIK {old_cik})")
        return company
