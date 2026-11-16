"""End-to-end number checks: build a workbook, calculate it, and compare with values worked out by hand.

These need LibreOffice to evaluate the formulas and are skipped where it is not installed.
"""
import pytest

from edgar_comps import recalc, workbook
from test_model import company

pytestmark = pytest.mark.skipif(recalc.find_soffice() is None, reason="LibreOffice not installed")

# The synthetic company in test_model.py, latest year:
# revenue 120, operating income 24, D&A 6, net income 15, cash 40, debt 70, 10 diluted shares.
PRICE = 10.0


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    target, peer = company(), company()
    peer.ticker = "PEER"
    path = tmp_path_factory.mktemp("wb") / "model.xlsx"
    workbook.build([target, peer], "TEST", path,
                   prices={"TEST": (PRICE, "2026-01-02", "test"), "PEER": (20.0, "2026-01-02", "test")},
                   market={"beta": (1.2, 60, "2025-12"), "risk_free": (0.04, "2026-01-02")})
    with recalc.calculated(path) as values:
        yield path, values


def test_no_formula_errors(built):
    path, _ = built
    assert recalc.formula_errors(path) == []


def test_comps_row_matches_hand_calculation(built):
    _, wb = built
    row = {header.value: cell.value for header, cell in zip(wb["Comps"][5], wb["Comps"][6])}
    assert row["Market cap"] == pytest.approx(100)              # 10 shares x $10
    assert row["Net debt"] == pytest.approx(30)                 # 70 debt - 40 cash
    assert row["Enterprise value"] == pytest.approx(130)
    assert row["EBITDA"] == pytest.approx(30)                   # 24 + 6
    assert row["EV / EBITDA"] == pytest.approx(130 / 30)
    assert row["P / E"] == pytest.approx(100 / 15)
    assert row["Revenue growth"] == pytest.approx(0.2)          # 100 -> 120


def test_peer_statistics_exclude_the_target(built):
    _, wb = built
    comps = wb["Comps"]
    peer_multiple = comps["N7"].value                           # PEER: (10 x 20 + 30) / 30
    assert peer_multiple == pytest.approx(230 / 30)
    assert comps["N10"].value == pytest.approx(peer_multiple)   # peer median = the only peer


def test_wacc_matches_capm_by_hand(built):
    _, wb = built
    wacc = wb["WACC"]
    cost_of_equity = 0.04 + 1.2 * 0.05
    after_tax_debt = (0.04 + 0.02) * (1 - wacc["B11"].value)
    equity_weight = 100 / (100 + 70)
    assert wacc["B8"].value == pytest.approx(cost_of_equity)
    assert wacc["B17"].value == pytest.approx(equity_weight * cost_of_equity + (1 - equity_weight) * after_tax_debt)
    assert wb["DCF"]["D4"].value == pytest.approx(wacc["B17"].value)        # the DCF uses it


def test_dcf_matches_an_independent_calculation(built):
    """Recompute the whole DCF in Python from the assumption cells and compare with Excel's answer."""
    _, wb = built
    dcf = wb["DCF"]
    rate, growth, tax, margin, da, reinvest, nwc, g1 = (dcf[f"D{r}"].value for r in range(4, 12))
    revenue, present_value = 120.0, 0.0
    for year in range(1, 6):
        previous = revenue
        revenue *= 1 + g1 + (growth - g1) * (year - 1) / 5
        cash_flow = revenue * margin * (1 - tax) + revenue * da - revenue * reinvest - (revenue - previous) * nwc
        present_value += cash_flow / (1 + rate) ** year
    terminal = cash_flow * (1 + growth) / (rate - growth) / (1 + rate) ** 5
    enterprise_value = present_value + terminal
    assert dcf["B34"].value == pytest.approx(enterprise_value)
    assert dcf["B39"].value == pytest.approx((enterprise_value - 30) / 10)  # less net debt, per share
    assert dcf["B52"].value == pytest.approx(0, abs=0.005)                  # sensitivity grid ties out


def test_memo_reads_as_text_with_the_right_numbers(built):
    _, wb = built
    memo = {row[0]: row[2] for row in wb["Memo"].iter_rows(values_only=True) if row[0] and row[2]}
    assert "$10.00 per share" in memo["Market value"]
    assert "4.3x EBITDA" in memo["Market value"]
    assert all(not str(text).startswith("=") for text in memo.values())    # calculated text, not formulas
