"""Export the Memo sheet as a Markdown file: python -m edgar_comps.memo output/imax_comps.xlsx

The memo text lives in Excel formulas, so this needs LibreOffice to calculate them.
Without LibreOffice, open the workbook in Excel and read the Memo sheet there.
"""
import sys
from pathlib import Path

from .recalc import calculated


def export(xlsx_path):
    xlsx_path = Path(xlsx_path)
    with calculated(xlsx_path) as workbook:
        ws = workbook["Memo"]
        lines = [f"# {ws['A1'].value}", ""]
        for a, b, c in ws.iter_rows(min_row=4, max_col=3, values_only=True):
            if a and str(a).startswith(("1.", "2.")):
                lines += [f"## {a[3:]}", ""]
            elif a and str(a).startswith("Workings"):
                break
            elif a and str(a).startswith("Audit result"):
                lines += [f"**{a}**", ""]
            elif a and c:
                lines.append(f"- **{a}{' [' + b + ']' if b else ''}:** {c}")
            elif not a:
                lines.append("")
    out = xlsx_path.with_name(xlsx_path.stem + "_memo.md")
    out.write_text("\n".join(lines).strip() + "\n")
    return out


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    print(f"Wrote {export(sys.argv[1])}")
