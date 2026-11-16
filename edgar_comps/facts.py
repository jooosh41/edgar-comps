"""Turn raw XBRL company facts into a clean table of annual line items.

Companies tag the same line with different XBRL concepts, so each line item
has a priority list of tags. For each fiscal year the first tag with a value
wins, and the tag used is recorded so every number can be traced.
"""
from dataclasses import dataclass, field
from datetime import date

ANNUAL_FORMS = {"10-K", "10-K/A", "10-KT"}

# (key, label, kind, unit, tag priority list). kind: flow = over the year, stock = at year end.
LINE_ITEMS = [
    ("revenue", "Revenue", "flow", "USD", [
        "Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax",
        "SalesRevenueNet", "RevenueFromContractWithCustomerIncludingAssessedTax"]),
    ("operating_income", "Operating income (reported)", "flow", "USD", ["OperatingIncomeLoss"]),
    ("pretax_income", "Pre-tax income", "flow", "USD", [
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments",
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesDomestic"]),
    ("interest_expense", "Interest expense", "flow", "USD", [
        "InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt",
        "InterestAndDebtExpense"]),
    # Positive = net interest income. Used for EBIT only when there is no interest expense tag.
    ("interest_net", "Net interest income / (expense) (fallback)", "flow", "USD", [
        "InterestIncomeExpenseNonoperatingNet", "InterestIncomeExpenseNet"]),
    ("tax_expense", "Income tax expense", "flow", "USD", ["IncomeTaxExpenseBenefit"]),
    ("net_income", "Net income", "flow", "USD", ["NetIncomeLoss", "ProfitLoss"]),
    ("d_and_a", "Depreciation & amortization", "flow", "USD", [
        "DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
        "DepreciationAmortizationAndAccretionNet"]),
    # Used only when the combined D&A row above is blank (see "D&A used" formula).
    ("depreciation_only", "Depreciation (fallback)", "flow", "USD", ["Depreciation"]),
    ("amortization_intangibles", "Amortization of intangibles (fallback)", "flow", "USD", [
        "AmortizationOfIntangibleAssets"]),
    ("cfo", "Cash from operations", "flow", "USD", [
        "NetCashProvidedByUsedInOperatingActivities",
        "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"]),
    ("capex", "Capital expenditures", "flow", "USD", [
        "PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"]),
    ("capex_intangibles", "Purchases of intangible assets", "flow", "USD", ["PaymentsToAcquireIntangibleAssets"]),
    ("investing_total", "Net cash from investing activities", "flow", "USD", ["NetCashProvidedByUsedInInvestingActivities"]),
    ("diluted_shares", "Diluted shares (weighted avg)", "flow", "shares", [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
        "WeightedAverageNumberOfShareOutstandingBasicAndDiluted",
        "WeightedAverageNumberOfSharesOutstandingBasic"]),
    ("cash", "Cash & equivalents", "stock", "USD", [
        "CashAndCashEquivalentsAtCarryingValue",
        "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"]),
    ("nci", "Non-controlling interest (book value)", "stock", "USD", ["MinorityInterest"]),
    ("nci_redeemable", "Redeemable non-controlling interest", "stock", "USD", [
        "RedeemableNoncontrollingInterestEquityCarryingAmount"]),
    ("debt_noncurrent", "Long-term debt, non-current", "stock", "USD", [
        "LongTermDebtNoncurrent", "LongTermDebtAndCapitalLeaseObligations"]),
    ("debt_current", "Debt, current portion", "stock", "USD", [
        "LongTermDebtCurrent", "DebtCurrent", "ShortTermBorrowings",
        "LongTermDebtAndCapitalLeaseObligationsCurrent"]),
    ("debt_total_reported", "Total debt (single reported tag)", "stock", "USD", [
        "LongTermDebt", "DebtAndCapitalLeaseObligations",
        "DebtLongtermAndShorttermCombinedAmount"]),
    # Used only when none of the three debt rows above has a value (see Total debt formula).
    ("debt_convertible", "Convertible debt (fallback)", "stock", "USD", [
        "ConvertibleDebt", "ConvertibleNotesPayable", "ConvertibleDebtNoncurrent"]),
    ("debt_revolver", "Credit facility borrowings (fallback)", "stock", "USD", [
        "LineOfCredit", "LongTermLineOfCredit"]),
]


@dataclass
class Line:
    key: str
    label: str
    unit: str
    values: list                      # one per fiscal year, None where missing
    tag: str = ""                     # tag behind the latest year's value
    accession: str = ""               # filing behind the latest year's value


@dataclass
class Company:
    ticker: str
    name: str
    cik: int
    year_ends: list                   # ISO dates, oldest first
    lines: dict = field(default_factory=dict)
    cover_shares: tuple = None                  # (count, as-of date) from the latest filing's cover page
    notes: list = field(default_factory=list)   # loader remarks, e.g. history taken from a predecessor entity

    def latest(self, key):
        return self.lines[key].values[-1]


def _days(start, end):
    return (date.fromisoformat(end) - date.fromisoformat(start)).days


def _annual_points(concept, unit, kind):
    """Map fiscal-year-end date -> (value, accession), keeping the latest filing."""
    best = {}
    for fact in concept.get("units", {}).get(unit, []):
        if kind == "flow":
            if fact.get("form") not in ANNUAL_FORMS or "start" not in fact:
                continue
            if not 350 <= _days(fact["start"], fact["end"]) <= 380:
                continue
        elif "start" in fact:
            continue
        end = fact["end"]
        if end not in best or fact["filed"] > best[end][2]:
            best[end] = (fact["val"], fact["accn"], fact["filed"])
    return {end: (v, accn) for end, (v, accn, _) in best.items()}


def build_company(ticker, name, cik, raw, years=5):
    gaap = raw.get("facts", {}).get("us-gaap", {})
    series = {}
    for key, label, kind, unit, tags in LINE_ITEMS:
        series[key] = [(tag, _annual_points(gaap[tag], unit, kind)) for tag in tags if tag in gaap]

    revenue_ends = sorted({end for _, pts in series["revenue"] for end in pts})
    # Keep fiscal years only: drop dates that are not ~1 year apart from the latest.
    year_ends = []
    for end in reversed(revenue_ends):
        if not year_ends or 350 <= _days(end, year_ends[-1]) <= 380:
            year_ends.append(end)
        if len(year_ends) == years:
            break
    year_ends.reverse()
    if not year_ends:
        raise ValueError(f"{ticker}: no annual revenue found in 10-K XBRL data")

    company = Company(ticker=ticker.upper(), name=name, cik=cik, year_ends=year_ends)
    for key, label, kind, unit, tags in LINE_ITEMS:
        line = Line(key=key, label=label, unit=unit, values=[])
        for end in year_ends:
            hit = next(((tag, pts[end]) for tag, pts in series[key] if end in pts), None)
            line.values.append(hit[1][0] if hit else None)
            if end == year_ends[-1] and hit:
                line.tag, line.accession = hit[0], hit[1][1]
        company.lines[key] = line

    cover = raw.get("facts", {}).get("dei", {}).get("EntityCommonStockSharesOutstanding", {}).get("units", {}).get("shares", [])
    if cover:
        latest = max(cover, key=lambda fact: (fact["end"], fact["filed"]))
        if latest["end"] >= year_ends[-1]:              # ignore stale counts (multi-class filers stop reporting one)
            company.cover_shares = (latest["val"], latest["end"])
    return company


# Items a company needs before its comps row or DCF can be trusted.
REQUIRED = ["revenue", "net_income", "cash"]
DEBT_STANDARD = ["debt_noncurrent", "debt_current", "debt_total_reported"]
DEBT_FALLBACK = ["debt_convertible", "debt_revolver"]


def _has(company, keys):
    return any(company.latest(k) is not None for k in keys)


def _ebit_derivable(company):
    has_interest = _has(company, ["interest_expense", "interest_net"])
    return company.latest("pretax_income") is not None and has_interest


def missing_items(company):
    """Labels of items the model cannot fill, even with fallbacks, in the latest fiscal year."""
    gaps = [company.lines[k].label for k in REQUIRED if company.latest(k) is None]
    if company.latest("diluted_shares") is None and company.cover_shares is None:
        gaps.append("Share count")
    if company.latest("operating_income") is None and not _ebit_derivable(company):
        gaps.append("Operating income")
    if not _has(company, ["d_and_a", "depreciation_only", "amortization_intangibles"]):
        gaps.append("Depreciation & amortization")
    if not _has(company, DEBT_STANDARD + DEBT_FALLBACK):
        gaps.append("Debt (no tag found; may be zero)")
    return gaps


def fallbacks_used(company):
    """Approximations standing in for a line the company does not tag in the standard way."""
    notes = list(company.notes)
    if company.latest("operating_income") is None and _ebit_derivable(company):
        if company.latest("interest_expense") is not None:
            notes.append("EBIT = pre-tax income + interest expense (company reports no operating income)")
        else:
            notes.append("EBIT = pre-tax income less net interest income (company reports no operating income)")
    if company.cover_shares is None and company.lines["diluted_shares"].tag and "Diluted" not in company.lines["diluted_shares"].tag:
        notes.append("Share count is basic, not diluted (no diluted tag)")
    if company.latest("d_and_a") is None and _has(company, ["depreciation_only", "amortization_intangibles"]):
        notes.append("D&A = depreciation + amortization of intangibles (no combined tag)")
    if not _has(company, DEBT_STANDARD) and _has(company, DEBT_FALLBACK):
        notes.append("Debt = convertible debt + credit facility (no standard debt tag)")
    return notes
