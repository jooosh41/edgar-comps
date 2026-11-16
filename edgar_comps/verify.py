"""Check a generated workbook: python -m edgar_comps.verify output/imax_comps.xlsx

Reports formula errors, the audit result, and whether the sensitivity grid ties to the DCF.
Exits with status 1 if anything is wrong, so it can gate a release.
"""
import sys

from .recalc import calculated, formula_errors


def verify(path):
    problems = [f"{sheet}!{cell} = {error}" for sheet, cell, error in formula_errors(path)]
    with calculated(path) as wb:
        tie_out = wb["DCF"]["B52"].value
        if not isinstance(tie_out, (int, float)) or abs(tie_out) > 0.005:
            problems.append(f"Sensitivity grid does not tie to the DCF (difference: {tie_out})")
        audit_line = next((row[0] for row in wb["Memo"].iter_rows(values_only=True)
                           if row[0] and str(row[0]).startswith("Audit result")), "Audit result not found")
    return problems, audit_line


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    found, audit = verify(sys.argv[1])
    print(audit)
    print("Formula errors and integrity problems:", len(found))
    for problem in found:
        print("  ", problem)
    sys.exit(1 if found else 0)
