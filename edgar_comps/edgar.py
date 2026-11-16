"""Download layer: ticker -> CIK -> XBRL company facts from SEC EDGAR.

Every response is cached (see paths.cache_dir) so repeated runs do not hit the SEC.
"""
import json
import os
import time
import requests

from . import paths

TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
MIN_SECONDS_BETWEEN_REQUESTS = 0.2  # SEC limit is 10 requests/second

_last_request = 0.0


def load_env():
    """Read KEY=VALUE lines from ./.env into the environment (no dependency needed)."""
    env_file = paths.env_file()
    if not env_file.exists():
        return
    for line in env_file.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            os.environ.setdefault(key.strip(), value.strip())


def _user_agent():
    load_env()
    ua = os.environ.get("SEC_USER_AGENT")
    if not ua:
        raise RuntimeError(
            "SEC_USER_AGENT is not set. The SEC blocks requests that do not declare a contact. Either put\n"
            "  SEC_USER_AGENT=edgar-comps your-email@example.com\n"
            "in a file named .env in this folder, or set it as an environment variable."
        )
    return ua


def _get_json(url, cache_name, max_age_hours=24):
    global _last_request
    path = paths.cache_dir() / cache_name
    if path.exists() and time.time() - path.stat().st_mtime < max_age_hours * 3600:
        return json.loads(path.read_text())
    wait = MIN_SECONDS_BETWEEN_REQUESTS - (time.time() - _last_request)
    if wait > 0:
        time.sleep(wait)
    response = requests.get(url, headers={"User-Agent": _user_agent()}, timeout=60)
    _last_request = time.time()
    response.raise_for_status()
    path.write_text(response.text)
    return response.json()


def lookup(ticker):
    """Return (cik, company_name) for a ticker, or raise KeyError."""
    table = _get_json(TICKER_URL, "company_tickers.json", max_age_hours=24 * 7)
    wanted = ticker.upper()
    for row in table.values():
        if row["ticker"].upper() == wanted:
            return int(row["cik_str"]), row["title"]
    raise KeyError(f"Ticker {ticker!r} not found in SEC ticker list")


def company_facts(cik):
    return _get_json(FACTS_URL.format(cik=cik), f"facts_{cik:010d}.json")
