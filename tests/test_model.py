"""Offline tests: no network. Run with `python -m pytest`."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from openpyxl import load_workbook

from edgar_comps import facts, workbook


def fact(val, end, start=None, form="10-K", filed="2025-02-01", accn="0000000000-25-000001"):
    row = {"val": val, "end": end, "form": form, "filed": filed, "accn": accn}
    if start:
        row["start"] = start
    return row


def raw_facts():
    usd = lambda rows: {"units": {"USD": rows}}
    return {"facts": {"us-gaap": {
        "SalesRevenueNet": usd([                                  # fallback tag, used only for 2023
            fact(80, "2023-12-31", "2023-01-01", filed="2024-02-01")]),
        "RevenueFromContractWithCustomerExcludingAssessedTax": usd([
            fact(100, "2024-12-31", "2024-01-01"),
            fact(30, "2024-12-31", "2024-10-01", form="10-K"),    # a quarter inside the 10-K: must be ignored
            fact(120, "2025-12-31", "2025-01-01", filed="2026-02-01"),
            fact(999, "2025-06-30", "2025-01-01", form="10-Q")]), # half-year: must be ignored
        "OperatingIncomeLoss": usd([
            fact(20, "2025-12-31", "2025-01-01", filed="2026-02-01"),
            fact(24, "2025-12-31", "2025-01-01", filed="2026-08-01", accn="restated")]),  # restatement wins
        "DepreciationDepletionAndAmortization": usd([fact(6, "2025-12-31", "2025-01-01")]),
        "NetIncomeLoss": usd([fact(15, "2025-12-31", "2025-01-01")]),
        "CashAndCashEquivalentsAtCarryingValue": usd([fact(40, "2025-12-31")]),
        "LongTermDebtNoncurrent": usd([fact(70, "2025-12-31")]),
        "WeightedAverageNumberOfDilutedSharesOutstanding": {"units": {"shares": [
            fact(10, "2025-12-31", "2025-01-01")]}},
    }}}


def company():
    return facts.build_company("test", "Test Co", 1, raw_facts())


def test_only_full_fiscal_years_are_kept():
    assert company().year_ends == ["2023-12-31", "2024-12-31", "2025-12-31"]


def test_fallback_tag_fills_older_year():
    assert company().lines["revenue"].values == [80, 100, 120]


def test_latest_filing_wins_and_is_traced():
    line = company().lines["operating_income"]
    assert line.values[-1] == 24 and line.accession == "restated"


def test_missing_items_are_reported():
    c = company()
    assert facts.missing_items(c) == []
    c.lines["operating_income"].values[-1] = None
    assert facts.missing_items(c) == ["Operating income"]


def test_workbook_uses_formulas_not_pasted_results(tmp_path):
    path = workbook.build([company()], "TEST", tmp_path / "t.xlsx")
    wb = load_workbook(path)
    assert wb.sheetnames == ["Inputs", "Memo", "Comps", "DCF", "WACC", "Checks", "D_TEST"]
    assert wb["D_TEST"]["F6"].value == 120                      # filing data is a value
    assert str(wb["D_TEST"][f"F{workbook.ROW['ebitda']}"].value).startswith("=")   # EBITDA is a formula
    assert "D_TEST" in wb["Comps"]["I6"].value                  # comps link to the data sheet
    assert wb["DCF"]["E17"].value == "=D17*(1+E18)"             # projections chain off history


def test_fallbacks_are_disclosed_not_silent():
    c = company()
    c.lines["operating_income"].values[-1] = None
    c.lines["pretax_income"].values[-1] = 18
    c.lines["interest_expense"].values[-1] = 2
    assert facts.missing_items(c) == []
    assert any("pre-tax income" in note for note in facts.fallbacks_used(c))


def test_memo_sentences_are_formulas(tmp_path):
    wb = load_workbook(workbook.build([company()], "TEST", tmp_path / "t.xlsx"))
    memo = wb["Memo"]
    assert str(memo["C5"].value).startswith("=IFERROR(")        # write-up is live text
    statuses = [memo.cell(row=r, column=2).value for r in range(16, 34)]
    assert any(str(v).startswith("=IF(") for v in statuses)     # audit checks are formulas
