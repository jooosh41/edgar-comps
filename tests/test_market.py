"""Market-data parsing, tested offline with canned responses in place of the network."""
import json
from datetime import datetime, timezone

import pytest

from edgar_comps import market


def chart(closes, start_year=2021):
    stamps = [int(datetime(start_year + i // 12, i % 12 + 1, 1, 12, tzinfo=timezone.utc).timestamp())
              for i in range(len(closes))]
    meta = {"currentTradingPeriod": {"regular": {"end": stamps[-1] + 4 * 3600}}}
    return json.dumps({"chart": {"result": [{"meta": meta, "timestamp": stamps,
                                             "indicators": {"adjclose": [{"adjclose": closes}],
                                                            "quote": [{"close": closes}]}}]}})


def test_price_uses_the_last_completed_close(monkeypatch):
    monkeypatch.setattr(market, "_cached", lambda url, name, headers=None: chart([10, 11, 12.5]))
    last_bar = json.loads(chart([10, 11, 12.5]))["chart"]["result"][0]["timestamp"][-1]
    after_close, _, source = market.price("ANY", now=last_bar + 5 * 3600)
    mid_session, _, _ = market.price("ANY", now=last_bar + 3600)
    assert after_close == 12.5 and "closing" in source
    assert mid_session == 11                                    # today's bar is a live quote: skip it


def test_beta_recovers_a_known_slope(monkeypatch):
    index_returns = [0.02, -0.01, 0.03, -0.02, 0.01] * 8                    # 40 months
    index, stock = [100.0], [50.0]
    for r in index_returns:
        index.append(index[-1] * (1 + r))
        stock.append(stock[-1] * (1 + 2 * r))                               # the stock moves twice as much
    series = {"idx": chart(index), "stock": chart(stock)}
    monkeypatch.setattr(market, "_cached",
                        lambda url, name, headers=None: series["idx" if "idx_" in name else "stock"])
    slope, months, _ = market.beta("ANY")
    assert slope == pytest.approx(2.0, abs=1e-6)
    assert months == 39                                                     # the unfinished last month is dropped


def test_beta_refuses_short_histories(monkeypatch):
    monkeypatch.setattr(market, "_cached", lambda url, name, headers=None: chart([1, 2, 3, 4, 5, 6]))
    with pytest.raises(ValueError):
        market.beta("NEW")


def test_risk_free_skips_blank_holidays(monkeypatch):
    csv_text = "observation_date,DGS10\n2026-11-10,4.98\n2026-11-11,\n"
    monkeypatch.setattr(market, "_cached", lambda url, name, headers=None: csv_text)
    assert market.risk_free() == (pytest.approx(0.0498), "2026-11-10")


def test_gather_reports_failures_instead_of_crashing(monkeypatch):
    def broken(url, name, headers=None):
        raise OSError("network down")
    monkeypatch.setattr(market, "_cached", broken)
    data = market.gather(["A", "B"], "A")
    assert data["prices"] == {} and data["beta"] is None and len(data["errors"]) == 4
