"""Market data the SEC does not publish: share prices, beta and the risk-free rate.

Prices and beta come from Yahoo Finance's chart endpoint, which is unofficial and can change
without notice. The risk-free rate is the 10-year Treasury yield from FRED (series DGS10).
If either source fails, the tool still runs: type the values into the yellow cells instead.
None of these requests carry the SEC contact email.
"""
import csv
import io
import json
import time
from datetime import datetime, timezone
from urllib.parse import quote

import requests

from .edgar import CACHE

YAHOO = "https://query1.finance.yahoo.com/v8/finance/chart/{symbol}?range={range}&interval={interval}"
FRED = "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10"
HEADERS = {"User-Agent": "Mozilla/5.0"}
MAX_AGE_SECONDS = 12 * 3600


def _cached(url, name, headers=HEADERS):
    CACHE.mkdir(parents=True, exist_ok=True)
    path = CACHE / name
    if path.exists() and time.time() - path.stat().st_mtime < MAX_AGE_SECONDS:
        return path.read_text()
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    path.write_text(response.text)
    return response.text


def _chart(symbol, range_, interval):
    url = YAHOO.format(symbol=quote(symbol), range=range_, interval=interval)
    name = f"yahoo_{symbol.replace('^', 'idx_')}_{range_}_{interval}.json"
    return json.loads(_cached(url, name))["chart"]["result"][0]


def _day(timestamp):
    return datetime.fromtimestamp(timestamp, tz=timezone.utc).date().isoformat()


def price(ticker, now=None):
    """Most recent closing price: (price, ISO date, source).

    While the market is open the last daily bar is a live quote, not a close, so it is skipped.
    A closing price makes a run reproducible: the same day always gives the same number.
    """
    result = _chart(ticker, "5d", "1d")
    closes = result["indicators"]["quote"][0]["close"]
    bars = [(stamp, close) for stamp, close in zip(result["timestamp"], closes) if close is not None]
    session_end = result["meta"].get("currentTradingPeriod", {}).get("regular", {}).get("end")
    now = time.time() if now is None else now
    if session_end and len(bars) > 1 and _day(bars[-1][0]) == _day(session_end) and now < session_end:
        bars = bars[:-1]
    stamp, close = bars[-1]
    return float(close), _day(stamp), "Yahoo Finance, closing price"


def _monthly_returns(symbol):
    result = _chart(symbol, "5y", "1mo")
    closes = result["indicators"]["adjclose"][0]["adjclose"]
    months = [_day(t)[:7] for t in result["timestamp"]]
    levels = {m: c for m, c in zip(months, closes) if c and m != months[-1]}   # drop the unfinished current month
    keys = sorted(levels)
    return {b: levels[b] / levels[a] - 1 for a, b in zip(keys, keys[1:])}


def beta(ticker, index="^GSPC"):
    """Slope of monthly returns against the S&P 500 over up to five years: (beta, observations, last month)."""
    stock, market = _monthly_returns(ticker), _monthly_returns(index)
    months = sorted(set(stock) & set(market))
    if len(months) < 24:
        raise ValueError(f"{ticker}: only {len(months)} months of price history, too few for a beta")
    xs, ys = [market[m] for m in months], [stock[m] for m in months]
    mean_x, mean_y = sum(xs) / len(xs), sum(ys) / len(ys)
    covariance = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    variance = sum((x - mean_x) ** 2 for x in xs)
    return covariance / variance, len(months), months[-1]


def risk_free():
    """Latest 10-year Treasury yield as a fraction: (rate, ISO date)."""
    rows = list(csv.reader(io.StringIO(_cached(FRED, "fred_DGS10.csv", headers={"User-Agent": "curl/8.0"}))))[1:]
    for day, value in reversed(rows):
        try:
            return float(value) / 100, day
        except ValueError:
            continue                                            # holidays are blank
    raise ValueError("No Treasury yield found in the FRED file")


def gather(tickers, target):
    """Everything the workbook needs, with failures reported instead of raised."""
    data = {"prices": {}, "beta": None, "risk_free": None, "errors": []}
    for ticker in tickers:
        try:
            data["prices"][ticker] = price(ticker)
        except Exception as error:
            data["errors"].append(f"price {ticker}: {error}")
    for key, call in (("beta", lambda: beta(target)), ("risk_free", risk_free)):
        try:
            data[key] = call()
        except Exception as error:
            data["errors"].append(f"{key}: {error}")
    return data
