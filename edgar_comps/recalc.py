"""Calculate a workbook's formulas with LibreOffice so Python can read the results.

openpyxl writes formulas but cannot evaluate them. Excel evaluates them when the file is opened;
this module does the same headlessly, for the memo export, the verify command and the tests.
"""
import shutil
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path

from openpyxl import load_workbook

CANDIDATES = ["soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def find_soffice():
    for name in CANDIDATES:
        found = shutil.which(name) or (name if Path(name).exists() else None)
        if found:
            return found
    return None


@contextmanager
def calculated(xlsx_path):
    """Yield an openpyxl workbook holding calculated values (not formulas) for xlsx_path."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice not found. Install it, or open the workbook in Excel instead.")
    xlsx_path = Path(xlsx_path)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, str(xlsx_path)],
                       check=True, capture_output=True, timeout=180)
        yield load_workbook(Path(tmp) / xlsx_path.name, data_only=True)


def formula_errors(xlsx_path):
    """Every formula cell whose calculated value is an Excel error: [(sheet, cell, error), ...]."""
    formulas = load_workbook(xlsx_path)
    errors = []
    with calculated(xlsx_path) as values:
        for sheet in formulas:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        value = values[sheet.title][cell.coordinate].value
                        if isinstance(value, str) and value.startswith(("#", "Err:")):
                            errors.append((sheet.title, cell.coordinate, value))
    return errors
