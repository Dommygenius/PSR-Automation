"""Write cell values into .xlsb/.xlsx via Excel COM — preserves all formatting, charts, pivots."""

from __future__ import annotations

import os
import shutil
import time
from datetime import date, datetime

import pandas as pd


def _cell_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.date()
    if isinstance(value, date):
        return value
    return value


def _import_win32():
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "The customer .xlsb workbook requires Excel + pywin32 on Windows.\n"
            "Install: pip install pywin32\n"
            "Excel must be installed (same machine as your other automations)."
        ) from exc
    return pythoncom, win32


def _open_excel_workbook(path: str, *, read_only: bool = False):
    pythoncom, win32 = _import_win32()
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(
        os.path.abspath(path),
        0,
        read_only,
    )
    return pythoncom, excel, wb


def _close_excel(pythoncom, excel, wb, *, save: bool = False) -> None:
    try:
        if save:
            wb.Save()
        wb.Close(SaveChanges=save)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


def copy_master_workbook(master_path: str, output_path: str) -> None:
    """Byte-for-byte copy of the master .xlsb — no openpyxl, no conversion."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    shutil.copy2(master_path, output_path)


def write_dataframe_values(
    ws,
    df: pd.DataFrame,
    start_row: int,
    start_col: int = 1,
    *,
    write_headers: bool = True,
    max_rows: int | None = None,
) -> int:
    """Set .Value only on the target block — no clear, no format changes."""
    if df.empty:
        return 0
    data = df.head(max_rows) if max_rows else df
    cells = 0
    data_start = start_row
    if write_headers:
        for j, name in enumerate(data.columns):
            ws.Cells(start_row, start_col + j).Value = str(name)
            cells += 1
        data_start = start_row + 1

    for i, row in enumerate(data.itertuples(index=False), start=0):
        excel_row = data_start + i
        for j, value in enumerate(row):
            ws.Cells(excel_row, start_col + j).Value = _cell_value(value)
            cells += 1
    return cells


def fill_xlsb_workbook(
    workbook_path: str,
    fills: list[dict],
    *,
    refresh_pivots: bool = True,
) -> None:
    """
    fills: list of {
        'sheet': str,
        'df': DataFrame,
        'start_row': int,
        'start_col': int (default 1),
        'write_headers': bool (default True),
        'max_rows': int|None,
    }
    """
    print(f"Opening via Excel COM (preserving all sheets/charts): {workbook_path}")
    pythoncom, excel, wb = _open_excel_workbook(workbook_path, read_only=False)
    try:
        for spec in fills:
            sheet_name = spec["sheet"]
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                print(f"WARNING: sheet '{sheet_name}' not found — skipped.")
                continue
            print(f"Writing values only -> '{sheet_name}' "
                  f"row {spec['start_row']} col {spec.get('start_col', 1)} "
                  f"({len(spec['df'])} rows)")
            n = write_dataframe_values(
                ws,
                spec["df"],
                spec["start_row"],
                spec.get("start_col", 1),
                write_headers=spec.get("write_headers", True),
                max_rows=spec.get("max_rows"),
            )
            print(f"  {n} cell value(s) set (formatting/charts untouched).")

        if refresh_pivots:
            print("Refreshing pivots / formulas (RefreshAll)…")
            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone(-1)
            time.sleep(2)

        _close_excel(pythoncom, excel, wb, save=True)
    except Exception:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()
        pythoncom.CoUninitialize()
        raise

    print(f"Saved: {workbook_path}")
