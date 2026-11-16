"""Write the Excel model. Filing data goes in as values; everything else is a live formula.

Colour convention (standard banking): blue = hardcoded input or filing data,
black = formula, green = link to another sheet, yellow fill = cell for you to fill in.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import evaluator
from .facts import LINE_ITEMS, fallbacks_used, missing_items

FONT = "Arial"
BLUE = Font(name=FONT, size=10, color="0000FF")
BLACK = Font(name=FONT, size=10, color="000000")
GREEN = Font(name=FONT, size=10, color="008000")
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=13, bold=True)
NOTE = Font(name=FONT, size=9, italic=True, color="595959")
HEAD_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
YELLOW = PatternFill("solid", fgColor="FFFF00")
TOP_RULE = Border(top=Side(style="thin"))

MM = '#,##0.0,,;(#,##0.0,,);"-"'        # dollars stored in full, shown in millions
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0"x"'
PRICE = '$#,##0.00'

YEAR_SLOTS = 5
FIRST_YEAR_COL = 2                       # column B
LATEST = get_column_letter(FIRST_YEAR_COL + YEAR_SLOTS - 1)   # column F holds the latest year
ITEM_ROW = {key: 6 + i for i, (key, *_rest) in enumerate(LINE_ITEMS)}
COVER_ROW = 6 + len(LINE_ITEMS)
DERIVED_HEADER_ROW = COVER_ROW + 1
DERIVED = ["ebit", "da", "shares", "nci_total", "reinvest", "cash_after_investing", "eff_tax", "total_debt", "net_debt", "ebitda", "fcf", "growth", "ebitda_margin", "ebit_margin", "da_pct", "capex_pct"]
ROW = dict(ITEM_ROW, cover_shares=COVER_ROW, **{key: DERIVED_HEADER_ROW + 1 + i for i, key in enumerate(DERIVED)})

PEER_FIRST_ROW = 6                       # first peer row on Inputs and Comps


def sheet_name(ticker):
    return f"D_{ticker}"[:31]


def _ref(ticker, key, col=LATEST):
    return f"'{sheet_name(ticker)}'!{col}{ROW[key]}"


def _header(ws, row, labels, widths=None):
    for i, label in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=i, value=label)
        cell.font, cell.fill = HEAD_FONT, HEAD_FILL
        cell.alignment = Alignment(horizontal="center" if i > 1 else "left", wrap_text=True)
    for i, width in enumerate(widths or [], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width


def _put(ws, ref, value, font=BLACK, fmt=None, fill=None):
    cell = ws[ref]
    cell.value, cell.font = value, font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    return cell


def _data_sheet(wb, company):
    ws = wb.create_sheet(sheet_name(company.ticker))
    _put(ws, "A1", f"{company.name} ({company.ticker}): annual financials", TITLE)
    _put(ws, "A2", "Dollars are stored in full and displayed in $ millions; shares in millions. "
                   "Blue = value pulled from 10-K XBRL data. Black = formula.", NOTE)
    _put(ws, "A3", "Source: SEC EDGAR, https://data.sec.gov/api/xbrl/companyfacts/"
                   f"CIK{company.cik:010d}.json. Where a year was restated, the latest filed value is used.", NOTE)
    _header(ws, 5, ["Fiscal year end"] + [""] * YEAR_SLOTS + ["XBRL tag (latest year)", "Filing accession (latest year)"],
            [34] + [13] * YEAR_SLOTS + [46, 24])
    offset = YEAR_SLOTS - len(company.year_ends)          # right-align so latest year is always column F
    cols = [get_column_letter(FIRST_YEAR_COL + offset + i) for i in range(len(company.year_ends))]
    for col, end in zip(cols, company.year_ends):
        ws[f"{col}5"] = end

    for key, label, _kind, unit, _tags in LINE_ITEMS:
        r, line = ROW[key], company.lines[key]
        _put(ws, f"A{r}", label + ("" if unit == "USD" else " (mm)"))
        for col, value in zip(cols, line.values):
            if value is not None:
                _put(ws, f"{col}{r}", value, BLUE, MM)
        _put(ws, f"G{r}", line.tag or "not found", NOTE)
        _put(ws, f"H{r}", line.accession, NOTE)

    if company.cover_shares:
        count, as_of = company.cover_shares
        _put(ws, f"A{COVER_ROW}", f"Shares outstanding, cover page as of {as_of} (mm)")
        _put(ws, f"{LATEST}{COVER_ROW}", count, BLUE, MM)
        _put(ws, f"G{COVER_ROW}", "dei:EntityCommonStockSharesOutstanding", NOTE)
    else:
        _put(ws, f"A{COVER_ROW}", "Shares outstanding, cover page (not reported)")

    _put(ws, f"A{DERIVED_HEADER_ROW}", "Derived (formulas)", BOLD)
    for key, (label, fmt) in DERIVED_LABELS.items():
        _put(ws, f"A{ROW[key]}", label)
        for i in range(YEAR_SLOTS):
            col = get_column_letter(FIRST_YEAR_COL + i)
            prev = get_column_letter(FIRST_YEAR_COL + i - 1) if i else None
            formula = _derived_formula(key, col, prev)
            if formula:
                _put(ws, f"{col}{ROW[key]}", formula, BLACK, fmt)
    ws.freeze_panes = "B6"
    return ws


DERIVED_LABELS = {
    "ebit": ("EBIT used (reported, else pre-tax income adjusted for interest)", MM),
    "da": ("D&A used (combined tag, else depreciation + amortization)", MM),
    "shares": ("Shares used (cover page, else diluted average)", MM),
    "nci_total": ("Non-controlling interest, total", MM),
    "reinvest": ("Reinvestment (capex + intangibles)", MM),
    "cash_after_investing": ("Cash from operations plus investing", MM),
    "eff_tax": ("Effective tax rate", PCT),
    "total_debt": ("Total debt", MM),
    "net_debt": ("Net debt (debt less cash)", MM),
    "ebitda": ("EBITDA (EBIT + D&A)", MM),
    "fcf": ("Free cash flow (CFO less reinvestment)", MM),
    "growth": ("Revenue growth", PCT),
    "ebitda_margin": ("EBITDA margin", PCT),
    "ebit_margin": ("Operating margin", PCT),
    "da_pct": ("D&A % of revenue", PCT),
    "capex_pct": ("Reinvestment % of revenue", PCT),
}
assert list(DERIVED_LABELS) == DERIVED, "DERIVED_LABELS must list the derived rows in sheet order"


def _derived_formula(key, col, prev):
    """Excel formula for one derived row in one year column (prev = the prior year's column, if any).

    Every formula returns "" when an input it needs is blank, so gaps stay visible as blanks
    instead of turning into zeros or errors further down the model.
    """
    def c(item):                                   # the cell holding `item` in this year's column
        return f"{col}{ROW[item]}"

    def span(first, last):                         # adjacent rows, e.g. the two debt components
        return f"{c(first)}:{c(last)}"

    def needs(items, expression):                  # blank unless every listed item has a value
        blanks = ",".join(f'{c(item)}=""' for item in items)
        return f'=IF(OR({blanks}),"",{expression})' if len(items) > 1 else f'=IF({blanks},"",{expression})'

    if key == "ebit":
        from_interest = (f'IF({c("interest_expense")}<>"",{c("pretax_income")}+ABS({c("interest_expense")}),'
                         f'IF({c("interest_net")}<>"",{c("pretax_income")}-{c("interest_net")},""))')
        return (f'=IF({c("operating_income")}<>"",{c("operating_income")},'
                f'IF({c("pretax_income")}="","",{from_interest}))')
    if key == "da":
        parts = span("depreciation_only", "amortization_intangibles")
        return f'=IF({c("d_and_a")}<>"",{c("d_and_a")},IF(COUNT({parts})>0,SUM({parts}),""))'
    if key == "shares":
        return (f'=IF({c("cover_shares")}<>"",{c("cover_shares")},'
                f'IF({c("diluted_shares")}="","",{c("diluted_shares")}))')
    if key == "nci_total":
        return needs(["revenue"], f'SUM({span("nci", "nci_redeemable")})')
    if key == "reinvest":
        return needs(["capex"], f'{c("capex")}+N({c("capex_intangibles")})')
    if key == "cash_after_investing":
        return needs(["cfo", "investing_total"], f'{c("cfo")}+{c("investing_total")}')
    if key == "eff_tax":
        rate = f'IF({c("pretax_income")}>0,{c("tax_expense")}/{c("pretax_income")},"")'
        return needs(["tax_expense", "pretax_income"], rate)
    if key == "total_debt":
        standard, fallback = span("debt_noncurrent", "debt_current"), span("debt_convertible", "debt_revolver")
        reported = c("debt_total_reported")
        return needs(["revenue"], f'IF(COUNT({standard})>0,SUM({standard}),IF({reported}<>"",{reported},'
                                  f'IF(COUNT({fallback})>0,SUM({fallback}),"")))')
    if key == "net_debt":
        return needs(["revenue"], f'N({c("total_debt")})-N({c("cash")})')
    if key == "ebitda":
        return needs(["ebit", "da"], f'{c("ebit")}+{c("da")}')
    if key == "fcf":
        return needs(["cfo", "reinvest"], f'{c("cfo")}-{c("reinvest")}')
    if key == "growth":
        if prev is None:
            return None
        return f'=IF(OR({prev}{ROW["revenue"]}="",{c("revenue")}=""),"",{c("revenue")}/{prev}{ROW["revenue"]}-1)'
    ratio_of = {"ebitda_margin": "ebitda", "ebit_margin": "ebit", "da_pct": "da", "capex_pct": "reinvest"}
    return _ratio(col, ratio_of[key])


def _ratio(col, key):
    num, den = f"{col}{ROW[key]}", f"{col}{ROW['revenue']}"
    return f'=IF(OR({num}="",{den}=""),"",{num}/{den})'


def _inputs_sheet(ws, companies, prices):
    _put(ws, "A1", "Inputs", TITLE)
    _put(ws, "A2", "Fill in the yellow cells. EDGAR has no market prices, so share prices are typed in by you "
                   "(or loaded with --prices). Valuation multiples stay blank until a price is entered.", NOTE)
    _put(ws, "A3", "Colour key: blue = hardcoded input or filing data; black = formula; green = link to another sheet.", NOTE)
    _header(ws, 5, ["Ticker", "Company", "Latest fiscal year end", "Share price ($)", "Price date", "Price source"],
            [10, 34, 20, 15, 14, 30])
    for i, c in enumerate(companies):
        r = PEER_FIRST_ROW + i
        _put(ws, f"A{r}", c.ticker, BOLD)
        _put(ws, f"B{r}", c.name)
        _put(ws, f"C{r}", c.year_ends[-1])
        price, when, source = prices.get(c.ticker, (None, None, None))
        _put(ws, f"D{r}", price, BLUE, PRICE, YELLOW)
        _put(ws, f"E{r}", when, BLUE, fill=YELLOW)
        _put(ws, f"F{r}", source, BLUE, fill=YELLOW)


def _comps_sheet(wb, companies):
    ws = wb.create_sheet("Comps")
    _put(ws, "A1", "Comparable companies: latest fiscal year", TITLE)
    _put(ws, "A2", "$ millions except per-share price. First row is the target; peer statistics exclude it. "
                   "Fiscal year ends differ by company, so rows are not the same 12 months. Enterprise value = market cap + net debt + non-controlling interest. NM = negative denominator.", NOTE)
    heads = ["Ticker", "Company", "FYE", "Price ($)", "Shares (mm)", "Market cap", "Net debt", "Enterprise value",
             "Revenue", "EBITDA", "EBIT", "Net income", "EV / Revenue", "EV / EBITDA", "EV / EBIT", "P / E",
             "Revenue growth", "EBITDA margin"]
    _header(ws, 5, heads, [9, 30, 12, 10, 12, 12, 12, 12, 12, 12, 12, 12, 10, 10, 10, 10, 10, 10])
    ws.row_dimensions[5].height = 30
    first, last = PEER_FIRST_ROW, PEER_FIRST_ROW + len(companies) - 1

    def link(t, key):
        ref = _ref(t, key)
        return f'=IF({ref}="","",{ref})'

    def multiple(num, den, r):
        return f'=IF(OR({num}{r}="",{den}{r}=""),"",IF({den}{r}>0,{num}{r}/{den}{r},"NM"))'

    for i, c in enumerate(companies):
        r, t = first + i, c.ticker
        _put(ws, f"A{r}", t, BOLD)
        _put(ws, f"B{r}", c.name)
        _put(ws, f"C{r}", c.year_ends[-1])
        _put(ws, f"D{r}", f'=IF(Inputs!D{r}="","",Inputs!D{r})', GREEN, PRICE)
        _put(ws, f"E{r}", link(t, "shares"), GREEN, MM)
        _put(ws, f"F{r}", f'=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})', BLACK, MM)
        _put(ws, f"G{r}", link(t, "net_debt"), GREEN, MM)
        _put(ws, f"H{r}", f'=IF(F{r}="","",F{r}+N(G{r})+N({_ref(t, "nci_total")}))', BLACK, MM)
        for col, key in (("I", "revenue"), ("J", "ebitda"), ("K", "ebit"), ("L", "net_income")):
            _put(ws, f"{col}{r}", link(t, key), GREEN, MM)
        _put(ws, f"M{r}", multiple("H", "I", r), BLACK, MULT)
        _put(ws, f"N{r}", multiple("H", "J", r), BLACK, MULT)
        _put(ws, f"O{r}", multiple("H", "K", r), BLACK, MULT)
        _put(ws, f"P{r}", multiple("F", "L", r), BLACK, MULT)
        _put(ws, f"Q{r}", link(t, "growth"), GREEN, PCT)
        _put(ws, f"R{r}", link(t, "ebitda_margin"), GREEN, PCT)

    peer_first = first + 1 if len(companies) > 1 else first     # statistics exclude the target (row 6)
    for j, (label, fn) in enumerate((("Peer mean", "AVERAGE"), ("Peer median", "MEDIAN"), ("Peer high", "MAX"), ("Peer low", "MIN"))):
        r = last + 2 + j
        _put(ws, f"A{r}", label, BOLD)
        for col, fmt in (("M", MULT), ("N", MULT), ("O", MULT), ("P", MULT), ("Q", PCT), ("R", PCT)):
            rng = f"{col}{peer_first}:{col}{last}"
            cell = _put(ws, f"{col}{r}", f'=IF(COUNT({rng})=0,"",{fn}({rng}))', BLACK, fmt)
            if j == 0:
                cell.border = TOP_RULE
    ws.freeze_panes = "C6"
    return ws


ASSUMPTION_ROWS = {"wacc": 4, "terminal_growth": 5, "tax_rate": 6, "ebit_margin": 7, "da_pct": 8,
                   "capex_pct": 9, "nwc_pct": 10, "growth_y1": 11}
WACC_RESULT = "WACC!$B$17"


def _dcf_sheet(wb, target, inputs_row, assumptions=None):
    assumptions = assumptions or {}
    ws = wb.create_sheet("DCF")
    t = target.ticker
    n = len(target.year_ends)
    _put(ws, "A1", f"Discounted cash flow: {target.name} ({t})", TITLE)
    _put(ws, "A2", "$ millions except per share. Yellow cells in the Base column are placeholders, not conclusions: "
                   "replace them with numbers you can defend. Type in Override to change any assumption; Used feeds the model.", NOTE)
    _header(ws, 3, ["Assumption", "Base", "Override", "Used", "Where the base comes from"], [36, 13, 13, 13, 13, 13, 13, 13, 13])

    first_col = get_column_letter(FIRST_YEAR_COL + YEAR_SLOTS - n)
    if n >= 2:
        cagr = f"=({_ref(t, 'revenue')}/{_ref(t, 'revenue', first_col)})^(1/{n - 1})-1"
        cagr_note = f"Revenue CAGR over the {n - 1} years of history on {sheet_name(t)}"
    else:
        cagr, cagr_note = 0, "Only one year of history: placeholder, set your own"
    rows = [
        (4, "Discount rate (WACC)", f"=IF(ISNUMBER({WACC_RESULT}),{WACC_RESULT},0.09)", "From the WACC sheet; 9% placeholder until it has a share price, beta and risk-free rate"),
        (5, "Terminal growth rate", 0.025, "Placeholder. Long-run nominal growth."),
        (6, "Tax rate on EBIT", "=" + _tax_base(t), "3-year average effective tax rate if between 10% and 40%, else 25%"),
        (7, "Operating (EBIT) margin", f"={_ref(t, 'ebit_margin')}", "Latest fiscal year, held flat"),
        (8, "D&A % of revenue", f'=IF({_ref(t, "da_pct")}="",0,{_ref(t, "da_pct")})', "Latest fiscal year, held flat; zero if the filing has no D&A tag"),
        (9, "Reinvestment % of revenue", f'=IF({_ref(t, "capex_pct")}="",0,{_ref(t, "capex_pct")})',
         "Capex plus purchased intangibles, latest fiscal year, held flat; zero if the filing has no capex tag"),
        (10, "Change in working capital % of revenue change", 0.0, "Placeholder. Zero assumes no working-capital drag."),
        (11, "Year 1 revenue growth", cagr, cagr_note + ". Fades in a straight line toward terminal growth."),
    ]
    for r, label, base, note in rows:
        _put(ws, f"A{r}", label)
        hard = not isinstance(base, str)
        _put(ws, f"B{r}", base, BLUE if hard else GREEN, PCT, YELLOW if hard else None)
        key = next(k for k, row in ASSUMPTION_ROWS.items() if row == r)
        value, why = assumptions.get(key, (None, ""))
        _put(ws, f"C{r}", value, BLUE, PCT, YELLOW)
        _put(ws, f"D{r}", f'=IF(C{r}="",B{r},C{r})', BLACK, PCT)
        _put(ws, f"E{r}", note + (f"  OVERRIDE: {why}" if value is not None and why else ""), NOTE)

    hist_cols = ["B", "C", "D"]
    src_cols = [get_column_letter(FIRST_YEAR_COL + YEAR_SLOTS - 3 + i) for i in range(3)]   # D, E, F on the data sheet
    proj = ["E", "F", "G", "H", "I"]
    _header(ws, 14, ["Free cash flow build"] + ["Actual"] * 3 + [f"Year {i}" for i in range(1, 6)])
    _put(ws, "A15", "Fiscal year end")
    _put(ws, "A16", "Years from valuation date")
    labels = {17: "Revenue", 18: "Revenue growth", 19: "EBIT", 20: "EBIT margin", 21: "Less: taxes on EBIT",
              22: "NOPAT", 23: "Plus: D&A", 24: "Less: reinvestment", 25: "Less: increase in working capital",
              26: "Unlevered free cash flow", 27: "Discount factor (year-end)", 28: "Present value of free cash flow",
              29: "Memo: actual cash from operations plus investing"}
    for r, label in labels.items():
        _put(ws, f"A{r}", label, BOLD if r in (17, 26, 28) else BLACK)

    def hist(key, col, sign=""):
        ref = _ref(t, key, col)
        return f'=IF({ref}="","",{sign}{ref})'

    for h, s in zip(hist_cols, src_cols):
        _put(ws, f"{h}15", f"='{sheet_name(t)}'!{s}5&\"\"", GREEN).alignment = Alignment(horizontal="right")
        _put(ws, f"{h}17", hist("revenue", s), GREEN, MM)
        _put(ws, f"{h}18", hist("growth", s), GREEN, PCT)
        _put(ws, f"{h}19", hist("ebit", s), GREEN, MM)
        _put(ws, f"{h}20", hist("ebit_margin", s), GREEN, PCT)
        _put(ws, f"{h}23", hist("da", s), GREEN, MM)
        _put(ws, f"{h}24", hist("reinvest", s, "-"), GREEN, MM)
        _put(ws, f"{h}29", hist("cash_after_investing", s), GREEN, MM)
    for i, c in enumerate(proj):
        p = "D" if i == 0 else proj[i - 1]
        _put(ws, f"{c}15", f"Year {i + 1}E").alignment = Alignment(horizontal="right")
        _put(ws, f"{c}16", i + 1, BLACK, "0")
        _put(ws, f"{c}17", f"={p}17*(1+{c}18)", BLACK, MM)
        _put(ws, f"{c}18", f"=$D$11+($D$5-$D$11)*({c}16-1)/5", BLACK, PCT)
        _put(ws, f"{c}19", f"={c}17*$D$7", BLACK, MM)
        _put(ws, f"{c}20", f"={c}19/{c}17", BLACK, PCT)
        _put(ws, f"{c}21", f"=-{c}19*$D$6", BLACK, MM)
        _put(ws, f"{c}22", f"={c}19+{c}21", BLACK, MM)
        _put(ws, f"{c}23", f"={c}17*$D$8", BLACK, MM)
        _put(ws, f"{c}24", f"=-{c}17*$D$9", BLACK, MM)
        _put(ws, f"{c}25", f"=-({c}17-{p}17)*$D$10", BLACK, MM)
        _put(ws, f"{c}26", f"=SUM({c}22:{c}25)", BOLD, MM).border = TOP_RULE
        _put(ws, f"{c}27", f"=1/(1+$D$4)^{c}16", BLACK, "0.000")
        _put(ws, f"{c}28", f"={c}26*{c}27", BOLD, MM)

    _put(ws, "A30", "Valuation", BOLD)
    val = [
        (31, "Sum of present values, years 1-5", "=SUM(E28:I28)", MM, BLACK),
        (32, "Terminal value (Gordon growth)", '=IF($D$4>$D$5,I26*(1+$D$5)/($D$4-$D$5),"WACC must exceed growth")', MM, BLACK),
        (33, "Present value of terminal value", '=IFERROR(B32*I27,"")', MM, BLACK),
        (34, "Enterprise value", '=IFERROR(B31+B33,"")', MM, BOLD),
        (35, "Terminal value % of enterprise value", '=IFERROR(B33/B34,"")', PCT, BLACK),
        (36, "Less: net debt and non-controlling interest", f"=N({_ref(t, 'net_debt')})+N({_ref(t, 'nci_total')})", MM, GREEN),
        (37, "Equity value", '=IFERROR(B34-B36,"")', MM, BOLD),
        (38, "Shares (mm)", f"={_ref(t, 'shares')}", MM, GREEN),
        (39, "Implied value per share", '=IFERROR(B37/B38,"")', PRICE, BOLD),
        (40, "Current share price", f'=IF(Inputs!D{inputs_row}="","",Inputs!D{inputs_row})', PRICE, GREEN),
        (41, "Implied upside / (downside)", '=IF(OR(B39="",B40=""),"",B39/B40-1)', PCT, BOLD),
    ]
    for r, label, formula, fmt, font in val:
        _put(ws, f"A{r}", label, BOLD if font is BOLD else BLACK)
        _put(ws, f"B{r}", formula, font, fmt)

    _put(ws, "A44", "Implied value per share: WACC (down) vs terminal growth (across)", BOLD)
    g_cols, g_steps = ["B", "C", "D", "E", "F"], [-0.01, -0.005, 0, 0.005, 0.01]
    _put(ws, "A45", "WACC \\ growth", NOTE)
    for c, step in zip(g_cols, g_steps):
        _put(ws, f"{c}45", f"=$D$5+({step})", BOLD, PCT)
    for k, step in enumerate([-0.01, -0.005, 0, 0.005, 0.01]):
        r = 46 + k
        _put(ws, f"A{r}", f"=$D$4+({step})", BOLD, PCT).alignment = Alignment(horizontal="right")
        for c in g_cols:
            _put(ws, f"{c}{r}",
                 f'=IFERROR(IF($A{r}>{c}$45,(SUMPRODUCT($E$26:$I$26/(1+$A{r})^$E$16:$I$16)'
                 f'+$I$26*(1+{c}$45)/($A{r}-{c}$45)/(1+$A{r})^$I$16-$B$36)/$B$38,""),"")', BLACK, PRICE)
    _put(ws, "A52", "Check: centre of grid less implied value per share (must be zero)", NOTE)
    _put(ws, "B52", '=IFERROR(D48-B39,"")', BLACK, '0.00;-0.00;"OK"')
    ws.freeze_panes = "B4"
    return ws


def _tax_base(ticker):
    cols = [get_column_letter(FIRST_YEAR_COL + YEAR_SLOTS - 3 + i) for i in range(3)]
    rng = f"'{sheet_name(ticker)}'!{cols[0]}{ROW['eff_tax']}:{cols[-1]}{ROW['eff_tax']}"
    return f"IFERROR(IF(AND(AVERAGE({rng})>=0.1,AVERAGE({rng})<=0.4),AVERAGE({rng}),0.25),0.25)"


def _wacc_sheet(wb, target, market):
    ws = wb.create_sheet("WACC")
    t = target.ticker
    rf, beta = market.get("risk_free"), market.get("beta")
    _put(ws, "A1", f"Discount rate build: {target.name} ({t})", TITLE)
    _put(ws, "A2", "CAPM cost of equity and a spread-based cost of debt, weighted at market values. Yellow cells are inputs; "
                   "run with --market to fill the first two automatically.", NOTE)
    _header(ws, 4, ["Input", "Value", "Source / note"], [40, 14, 110])
    rows = [
        (5, "Risk-free rate", rf[0] if rf else None, PCT, "input",
         f"10-year US Treasury yield, FRED series DGS10, {rf[1]}" if rf else "Not fetched. Enter the 10-year Treasury yield."),
        (6, "Beta", beta[0] if beta else None, "0.00", "input",
         f"Slope of {beta[1]} monthly returns against the S&P 500 through {beta[2]}, Yahoo Finance prices" if beta
         else "Not fetched. Enter a beta."),
        (7, "Equity risk premium", 0.05, PCT, "input", "Assumption. The common range is 4% to 6%."),
        (8, "Cost of equity", '=IF(COUNT(B5:B7)=3,B5+B6*B7,"")', PCT, "calc", "Risk-free rate + beta x equity risk premium"),
        (9, "Credit spread over risk-free", 0.02, PCT, "input", "Assumption. Replace with the yield on the company's own debt if you can find it."),
        (10, "Pre-tax cost of debt", '=IF(ISNUMBER(B5),B5+B9,"")', PCT, "calc", "Risk-free rate + credit spread"),
        (11, "Tax rate", "=DCF!D6", PCT, "link", "From the DCF sheet"),
        (12, "After-tax cost of debt", '=IF(ISNUMBER(B10),B10*(1-B11),"")', PCT, "calc", ""),
        (13, "Market value of equity ($mm)", f'=IF(Comps!F{PEER_FIRST_ROW}="","",Comps!F{PEER_FIRST_ROW})', MM, "link", "Share price x shares, from Comps"),
        (14, "Debt ($mm)", f"=N({_ref(t, 'total_debt')})", MM, "link", "Book value of debt, latest fiscal year"),
        (15, "Equity weight", '=IF(ISNUMBER(B13),B13/(B13+B14),"")', PCT, "calc", ""),
        (16, "Debt weight", '=IF(ISNUMBER(B15),1-B15,"")', PCT, "calc", ""),
        (17, "WACC", '=IF(AND(ISNUMBER(B8),ISNUMBER(B12),ISNUMBER(B15)),B15*B8+B16*B12,"")', PCT, "calc", "Feeds the DCF. Blank until a share price, beta and risk-free rate exist."),
        (19, "Reference: interest expense / debt",
         f'=IFERROR(ABS({_ref(t, "interest_expense")})/{_ref(t, "total_debt")},"")', PCT, "link",
         "What the company actually paid. Convertible debt carries a low coupon, so this can understate the true cost."),
    ]
    for r, label, value, fmt, kind, note in rows:
        _put(ws, f"A{r}", label, BOLD if r == 17 else BLACK)
        font = {"input": BLUE, "calc": BOLD if r == 17 else BLACK, "link": GREEN}[kind]
        _put(ws, f"B{r}", value, font, fmt, YELLOW if kind == "input" else None)
        _put(ws, f"C{r}", note, NOTE)
    return ws


def _checks_sheet(wb, companies):
    ws = wb.create_sheet("Checks")
    _put(ws, "A1", "Data checks", TITLE)
    _put(ws, "A2", "XBRL tagging is inconsistent across companies. Anything listed as missing was not found under a "
                   "standard tag: read the 10-K and fix the row on that company's data sheet before trusting its multiples.", NOTE)
    _header(ws, 4, ["Ticker", "Years of history", "Missing in latest fiscal year", "Debt: components less single reported tag ($mm)",
                    "Approximations in use"], [10, 16, 50, 34, 110])
    for i, c in enumerate(companies):
        r = 5 + i
        gaps = missing_items(c)
        _put(ws, f"A{r}", c.ticker, BOLD)
        _put(ws, f"B{r}", len(c.year_ends))
        _put(ws, f"C{r}", "; ".join(gaps) if gaps else "None", Font(name=FONT, size=10, color="C00000" if gaps else "000000"))
        _put(ws, f"E{r}", "; ".join(fallbacks_used(c)) or "None")
        s = f"'{sheet_name(c.ticker)}'!{LATEST}"
        _put(ws, f"D{r}", f'=IF(OR(COUNT({s}{ROW["debt_noncurrent"]}:{s[len(s) - 1:]}{ROW["debt_current"]})=0,{s}{ROW["debt_total_reported"]}=""),'
                          f'"n/a",{s}{ROW["total_debt"]}-{s}{ROW["debt_total_reported"]})', GREEN, MM)
    return ws


def build(companies, target_ticker, path, prices=None, market=None, assumptions=None):
    """companies: list of facts.Company. prices: {ticker: (price, date, source)}.
    market: output of market.gather(). assumptions: {key: (value, note)} for the DCF Override column."""
    market = market or {}
    prices = dict(market.get("prices", {}), **(prices or {}))
    target = next(c for c in companies if c.ticker == target_ticker.upper())
    companies = [target] + [c for c in companies if c is not target]      # target always sits on row 6
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    _inputs_sheet(ws, companies, prices)
    _comps_sheet(wb, companies)
    _dcf_sheet(wb, target, PEER_FIRST_ROW, assumptions)
    _wacc_sheet(wb, target, market)
    _checks_sheet(wb, companies)
    for c in companies:
        _data_sheet(wb, c)
    last_row = PEER_FIRST_ROW + len(companies) - 1
    evaluator.build_sheet(wb, target, companies[1:], median_row=last_row + 3, last_peer_row=last_row)
    wb.save(path)
    return path
