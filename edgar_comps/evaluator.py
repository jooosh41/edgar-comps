"""The Memo sheet: a rule-based write-up of the valuation and an audit of the model.

Every sentence is an Excel formula, so the text rewrites itself when a price or an
assumption changes. No sentence is free-written: each one is a fixed rule over cells
on the Comps and DCF sheets, which means every claim can be traced to a number.
"""
from datetime import date

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill

from .facts import fallbacks_used, missing_items

FONT = "Arial"
BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=13, bold=True)
NOTE = Font(name=FONT, size=9, italic=True, color="595959")
HEAD_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
WRAP = Alignment(wrap_text=True, vertical="top")

T = 6  # the target is always the first company row on Comps and Inputs


def _mm(ref):
    return f'TEXT({ref}/1000000,"#,##0;(#,##0)")'


def _usd(ref):
    return f'TEXT({ref},"#,##0.00")'


def _pct(ref, fmt="0.0%"):
    return f'TEXT({ref},"{fmt}")'


def _x(ref):
    return f'TEXT({ref},"0.0")&"x"'


def write_up(target, median_row, work):
    """(topic, formula) pairs for the valuation summary. work = cell refs of the workings block."""
    c, t, name = "Comps!", target.ticker, target.name.replace('"', "")
    q, r, n = f"{c}Q{T}", f"{c}R{T}", f"{c}N{T}"
    q_med, r_med, n_med = f"{c}Q{median_row}", f"{c}R{median_row}", f"{c}N{median_row}"
    price, comps_ps, dcf_ps = work["price"], work["comps_ps"], work["dcf_ps"]
    return [
        ("Business",
         f'=IFERROR("{name} generated revenue of $"&{_mm(f"{c}I{T}")}&" million in the fiscal year ended {target.year_ends[-1]}, "'
         f'&IF({q}="","",IF({q}>=0,"up ","down ")&{_pct(f"ABS({q})")}&" on the prior year, ")'
         f'&"with EBITDA of $"&{_mm(f"{c}J{T}")}&" million ("&{_pct(r)}&" margin) and EBIT of $"&{_mm(f"{c}K{T}")}&" million.",'
         f'"Not enough data for a business summary: see the Checks sheet.")'),
        ("Against peers",
         f'=IFERROR("Its EBITDA margin is "&TEXT(ABS({r}-{r_med})*100,"0.0")&" points "&IF({r}>={r_med},"above","below")'
         f'&" the peer median of "&{_pct(r_med)}&", and its revenue growth of "&{_pct(q)}&" compares with a peer median of "&{_pct(q_med)}&".",'
         f'"Peer comparison unavailable: a margin or growth figure is missing.")'),
        ("Leverage",
         f'=IFERROR(IF({c}G{T}<0,"The company holds net cash of $"&{_mm(f"-{c}G{T}")}&" million.",'
         f'"Net debt is $"&{_mm(f"{c}G{T}")}&" million"&IF({c}J{T}>0,", or "&{_x(f"{c}G{T}/{c}J{T}")}&" EBITDA.",". EBITDA is not positive, so a leverage multiple is not meaningful.")),'
         f'"Leverage unavailable.")'),
        ("Market value",
         f'=IF({price}="","No share price is entered for {t} on the Inputs sheet, so its trading multiples are blank.",'
         f'IFERROR("At $"&{_usd(price)}&" per share the market values {t} at an enterprise value of $"&{_mm(f"{c}H{T}")}&" million: "&{_x(f"{c}M{T}")}&" revenue and "'
         f'&IF(ISNUMBER({n}),{_x(n)}&" EBITDA","a non-meaningful EBITDA multiple")'
         f'&IF(AND(ISNUMBER({n}),ISNUMBER({n_med})),", a "&{_pct(f"ABS({n}/{n_med}-1)", "0%")}&IF({n}>={n_med}," premium"," discount")&" to the peer median of "&{_x(n_med)},"")&".",'
         f'"Trading multiples unavailable."))'),
        ("Value from comps",
         f'=IF(ISNUMBER({comps_ps}),"Applying the peer median EV/EBITDA of "&{_x(n_med)}&" to EBITDA of $"&{_mm(f"{c}J{T}")}&" million implies an enterprise value of $"'
         f'&{_mm(f"{n_med}*{c}J{T}")}&" million, or $"&{_usd(comps_ps)}&" per share.",'
         f'"No comps-implied value yet: it needs peer share prices on the Inputs sheet and positive EBITDA.")'),
        ("Value from DCF",
         f'=IFERROR("The DCF, at a "&{_pct("DCF!D4")}&" discount rate and "&{_pct("DCF!D5")}&" terminal growth, implies an enterprise value of $"&{_mm("DCF!B34")}'
         f'&" million and $"&{_usd(dcf_ps)}&" per share"&IF({price}="","",", "&{_pct("ABS(DCF!B41)", "0%")}&IF(DCF!B41>=0," above"," below")&" the current price")'
         f'&". Across the sensitivity grid, value per share runs from $"&{_usd("MIN(DCF!B46:F50)")}&" to $"&{_usd("MAX(DCF!B46:F50)")}&".",'
         f'"DCF unavailable: check the assumptions on the DCF sheet.")'),
        ("Reading",
         f'=IF({price}="","No reading without a share price for {t}.",IFERROR(IF(NOT(ISNUMBER({comps_ps})),'
         f'"Only the DCF is available. It puts value "&{_pct(f"ABS({dcf_ps}/{price}-1)", "0%")}&IF({dcf_ps}>={price}," above"," below")&" the current price.",'
         f'IF(AND({comps_ps}>{price}*1.1,{dcf_ps}>{price}*1.1),"Both methods put value more than 10% above the current price.",'
         f'IF(AND({comps_ps}<{price}*0.9,{dcf_ps}<{price}*0.9),"Both methods put value more than 10% below the current price.",'
         f'IF(AND(ABS({comps_ps}/{price}-1)<=0.1,ABS({dcf_ps}/{price}-1)<=0.1),"Both methods land within 10% of the current price.",'
         f'"The two methods disagree: comps imply $"&{_usd(comps_ps)}&" and the DCF $"&{_usd(dcf_ps)}&", against a price of $"&{_usd(price)}&". Resolve the gap before drawing a conclusion."))))'
         f'&" This is the model\'s output under its stated assumptions, not a recommendation.","No reading: a value is missing."))'),
    ]


def audit(target, peers, median_row, last_peer_row):
    """(check, status formula or text, finding formula or text) triples."""
    t = target.ticker
    n_med = f"Comps!N{median_row}"
    exit_mult = "DCF!B32/(DCF!I19+DCF!I23)"
    open_calls = '(IF(AND(DCF!C4="",NOT(ISNUMBER(WACC!$B$17))),1,0)+IF(DCF!C5="",1,0)+IF(DCF!C10="",1,0))'
    hist_margin = "AVERAGE(DCF!B20:D20)"
    target_issues = missing_items(target) + fallbacks_used(target)
    weak_peers = [p.ticker for p in peers if missing_items(p) or fallbacks_used(p)]
    end = date.fromisoformat(target.year_ends[-1])
    off_cycle = [f"{p.ticker} ({p.year_ends[-1]})" for p in peers
                 if abs((date.fromisoformat(p.year_ends[-1]) - end).days) > 92]
    return [
        ("Judgment assumptions",
         f'=IF({open_calls}>0,"WARNING","OK")',
         f'=IF({open_calls}>0,{open_calls}&" of 3 judgment assumptions (discount rate, terminal growth, working capital) are still tool defaults. '
         f'Set them in the Override column on the DCF sheet; until then the DCF value is not defensible.",'
         f'"Discount rate, terminal growth and working capital have all been set deliberately.")'),
        ("Discount rate source",
         '=IF(DCF!C4<>"","INFO",IF(ISNUMBER(WACC!$B$17),"OK","WARNING"))',
         f'=IF(DCF!C4<>"","The discount rate is a manual override of "&{_pct("DCF!C4")}&".",IF(ISNUMBER(WACC!$B$17),"The discount rate of "&{_pct("WACC!$B$17")}'
         f'&" is built on the WACC sheet: risk-free "&{_pct("WACC!$B$5")}&", beta "&TEXT(WACC!$B$6,"0.00")&", equity risk premium "&{_pct("WACC!$B$7")}&", equity weight "&{_pct("WACC!$B$15", "0%")}&".",'
         f'"The discount rate is the 9% placeholder: the WACC sheet lacks a share price, beta or risk-free rate."))'),
        ("Beta",
         '=IF(DCF!C4<>"","INFO",IF(NOT(ISNUMBER(WACC!$B$6)),"INFO",IF(OR(WACC!$B$6<0.6,WACC!$B$6>2),"WARNING","OK")))',
         '=IF(NOT(ISNUMBER(WACC!$B$6)),"No beta on the WACC sheet.","Beta is "&TEXT(WACC!$B$6,"0.00")&"."&IF(WACC!$B$6<0.6," Below 0.6 is unusually low for an operating company and pulls '
         'the discount rate down. A beta measured on one stock is noisy: compare it with the peers\' betas before relying on it.",IF(WACC!$B$6>2," Above 2 is unusually high: check the price history for a one-off shock.","")))'),
        ("Cash flow reconciliation",
         '=IFERROR(IF(OR(DCF!D29<=0,DCF!E26>1.3*DCF!D29),"WARNING","OK"),"n/a")',
         f'=IFERROR("The model projects year-1 free cash flow of $"&{_mm("DCF!E26")}&" million. Last year the company actually generated $"&{_mm("DCF!D29")}'
         f'&" million of cash from operations after all investing."&IF(OR(DCF!D29<=0,DCF!E26>1.3*DCF!D29)," The projection is more than 30% above what the business has delivered: '
         f'look for spending the model leaves out (investing lines beyond capex, working capital, acquisitions).",""),"Not available: operating or investing cash flow is missing.")'),
        ("Terminal value weight",
         '=IFERROR(IF(DCF!B35>0.75,"WARNING","OK"),"n/a")',
         f'=IFERROR("Terminal value is "&{_pct("DCF!B35", "0%")}&" of enterprise value."&IF(DCF!B35>0.75," Above 75%, the valuation rests mostly on what happens after year 5, '
         f'which is the least knowable part."," That is within the usual range."),"Terminal value not computed.")'),
        ("Discount rate vs growth",
         '=IFERROR(IF(OR(DCF!D4-DCF!D5<0.04,DCF!D5>0.04),"WARNING","OK"),"n/a")',
         '=IFERROR("The discount rate exceeds terminal growth by "&TEXT((DCF!D4-DCF!D5)*100,"0.0")&" points."&IF(DCF!D4-DCF!D5<0.04," Below 4 points, small changes in either '
         'input swing the value sharply.","")&IF(DCF!D5>0.04," Terminal growth above 4% implies the company outgrows the economy forever.",""),"n/a")'),
        ("Implied exit multiple",
         f'=IFERROR(IF(ISNUMBER({n_med}),IF({exit_mult}>1.5*{n_med},"WARNING","OK"),"INFO"),"n/a")',
         f'=IFERROR("The terminal value equals "&{_x(exit_mult)}&" year-5 EBITDA"&IF(ISNUMBER({n_med}),", against a peer median trading multiple of "&{_x(n_med)}'
         f'&IF({exit_mult}>1.5*{n_med},". The DCF assumes a far richer multiple than the market pays today.","."),". Enter peer prices to compare it with trading multiples."),"n/a")'),
        ("Margin assumption",
         f'=IFERROR(IF(DCF!D7>{hist_margin}+0.03,"WARNING","OK"),"n/a")',
         f'=IFERROR("Projected EBIT margin is "&{_pct("DCF!D7")}&" against a 3-year average of "&{_pct(hist_margin)}&"."&IF(DCF!D7>{hist_margin}+0.03,'
         f'" The model holds the latest year flat, and the latest year is well above the average: you are assuming a peak margin lasts forever.",""),"n/a")'),
        ("Growth assumption",
         '=IFERROR(IF(DCF!D11>0.15,"WARNING","OK"),"n/a")',
         f'=IFERROR("Year-1 revenue growth is "&{_pct("DCF!D11")}&", fading to "&{_pct("DCF!I18")}&" by year 5."&IF(DCF!D11>0.15," Above 15% deserves a specific justification; '
         f'a historical average can be inflated by a weak starting year.",""),"n/a")'),
        ("Reinvestment",
         '=IFERROR(IF(OR(DCF!D9<=0,DCF!D8>2*DCF!D9),"WARNING","OK"),"n/a")',
         f'=IFERROR("D&A is "&{_pct("DCF!D8")}&" of revenue and reinvestment is "&{_pct("DCF!D9")}&"."&IF(DCF!D9<=0," The model reinvests nothing: no capex was found for this company, so free cash flow is overstated. Set capex_pct yourself.",IF(DCF!D8>2*DCF!D9," D&A is more than twice reinvestment, so the model adds back far more '
         f'than it reinvests. Either the business is shrinking its asset base or spending sits outside the capex line (content, leased equipment, acquisitions). Free cash flow is probably overstated.","")),"n/a")'),
        ("DCF vs market",
         '=IF(DCF!B40="","INFO",IFERROR(IF(ABS(DCF!B41)>0.5,"WARNING","OK"),"n/a"))',
         f'=IF(DCF!B40="","No share price for {t}, so the DCF cannot be compared with the market.",IFERROR("The DCF value is "&{_pct("ABS(DCF!B41)", "0%")}&IF(DCF!B41>=0," above"," below")'
         f'&" the share price."&IF(ABS(DCF!B41)>0.5," A gap over 50% more often means an assumption is wrong than that the market is.",""),"n/a"))'),
        ("Profitability",
         f'=IFERROR(IF(Comps!K{T}<=0,"WARNING","OK"),"n/a")',
         f'=IFERROR(IF(Comps!K{T}<=0,"EBIT is not positive. A DCF that holds this margin flat projects losses forever, and EBITDA multiples may be non-meaningful.",'
         f'"EBIT is positive, so margin-based projections are mechanically valid."),"EBIT is missing.")'),
        ("Share prices",
         f'=IF(COUNTBLANK(Inputs!D{T}:D{last_peer_row})>0,"WARNING","OK")',
         f'=IF(COUNTBLANK(Inputs!D{T}:D{last_peer_row})>0,COUNTBLANK(Inputs!D{T}:D{last_peer_row})&" of {len(peers) + 1} share prices are missing on the Inputs sheet. '
         f'Multiples for those companies are blank and drop out of the peer statistics.","All share prices are entered. Check the price dates are close together.")'),
        (f"{t} data quality",
         "WARNING" if target_issues else "OK",
         ("Approximations or gaps in the target's data: " + "; ".join(target_issues) + ". Verify against the 10-K.")
         if target_issues else "All target lines came from standard tags with no approximation."),
        ("Peer data quality",
         "WARNING" if weak_peers else "OK",
         ("Peers with a gap or approximation (see Checks sheet): " + ", ".join(weak_peers) + ".")
         if weak_peers else "All peers loaded without gaps or approximations."),
        ("Fiscal year alignment",
         "WARNING" if off_cycle else "OK",
         ("Peers whose fiscal year ends more than 3 months from the target's: " + ", ".join(off_cycle) + ". Their figures cover a different 12 months.")
         if off_cycle else "All peers' fiscal years end within 3 months of the target's."),
        ("Share count",
         "INFO",
         (f"Market cap uses the share count on the cover of the latest filing ({target.cover_shares[1]}). It excludes dilution from options and convertible debt."
          if target.cover_shares else
          "Market cap uses the fiscal-year weighted-average diluted share count, not today's count. Buybacks or issuance since then are not reflected.")),
        ("Formula integrity",
         '=IFERROR(IF(ABS(DCF!B52)<0.005,"OK","WARNING"),"n/a")',
         '=IFERROR(IF(ABS(DCF!B52)<0.005,"The centre of the sensitivity grid ties to the DCF value per share.","The sensitivity grid no longer ties to the DCF: a formula has been overwritten."),"n/a")'),
    ]


def build_sheet(wb, target, peers, median_row, last_peer_row):
    ws = wb.create_sheet("Memo", 1)
    for col, width in (("A", 26), ("B", 12), ("C", 135)):
        ws.column_dimensions[col].width = width
    ws["A1"].value, ws["A1"].font = f"Valuation memo: {target.name} ({target.ticker})", TITLE
    ws["A2"].value = ("Generated by fixed rules from the Comps and DCF sheets. Every sentence is a formula and updates when "
                      "you change a price or an assumption.")
    ws["A2"].font = NOTE

    def header(row, labels):
        for i, label in enumerate(labels, start=1):
            cell = ws.cell(row=row, column=i, value=label)
            cell.font, cell.fill = HEAD_FONT, HEAD_FILL

    audit_rows = audit(target, peers, median_row, last_peer_row)
    summary_first = 5
    audit_first = summary_first + 7 + 4          # 7 summary rows, then a gap and two header rows
    work_first = audit_first + len(audit_rows) + 2
    work = {"price": f"$C${work_first + 1}", "comps_ps": f"$C${work_first + 2}", "dcf_ps": f"$C${work_first + 3}"}

    header(4, ["1. Valuation summary", "", ""])
    for i, (topic, formula) in enumerate(write_up(target, median_row, work)):
        r = summary_first + i
        ws.cell(row=r, column=1, value=topic).font = BOLD
        cell = ws.cell(row=r, column=3, value=formula)
        cell.font, cell.alignment = BODY, WRAP
        ws.row_dimensions[r].height = 42

    last_audit = audit_first + len(audit_rows) - 1
    ws.cell(row=audit_first - 2, column=1,
            value=f'="Audit result: "&COUNTIF(B{audit_first}:B{last_audit},"WARNING")&" warnings across {len(audit_rows)} checks."').font = BOLD
    header(audit_first - 1, ["2. Model audit", "Status", "Finding"])
    for i, (check, status, finding) in enumerate(audit_rows):
        r = audit_first + i
        ws.cell(row=r, column=1, value=check).font = BOLD
        ws.cell(row=r, column=2, value=status).font = BODY
        cell = ws.cell(row=r, column=3, value=finding)
        cell.font, cell.alignment = BODY, WRAP
        ws.row_dimensions[r].height = 42
    rng = f"B{audit_first}:B{last_audit}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"WARNING"'], font=Font(name=FONT, bold=True, color="C00000")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'], font=Font(name=FONT, color="008000")))

    header(work_first, ["Workings (used by the sentences above)", "", ""])
    workings = [
        ("Current share price", f'=IF(Comps!D{T}="","",Comps!D{T})', "$#,##0.00"),
        ("Comps-implied value per share",
         f'=IFERROR(IF(AND(ISNUMBER(Comps!N{median_row}),Comps!J{T}>0),(Comps!N{median_row}*Comps!J{T}-N(DCF!$B$36))/Comps!E{T},""),"")', "$#,##0.00"),
        ("DCF value per share", '=IF(DCF!B39="","",DCF!B39)', "$#,##0.00"),
    ]
    for i, (label, formula, fmt) in enumerate(workings):
        r = work_first + 1 + i
        ws.cell(row=r, column=1, value=label).font = BODY
        cell = ws.cell(row=r, column=3, value=formula)
        cell.font, cell.number_format, cell.alignment = BODY, fmt, Alignment(horizontal="left")
    return ws
