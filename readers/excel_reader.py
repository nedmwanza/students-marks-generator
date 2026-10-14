from pathlib import Path
import openpyxl


def load_workbook(path: str | Path):
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Please select an .xlsx or .xlsm Excel workbook.")

    # data_only=True reads saved formula results rather than formula text.
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=False)
    if not workbook.worksheets:
        raise ValueError("The workbook does not contain a worksheet.")
    return workbook
