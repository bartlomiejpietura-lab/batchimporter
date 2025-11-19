#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_refresh_all_v2_1.py
- Opens an .xlsx via COM, RefreshAll, saves, quits.
- Retries if Excel is busy or file is locked.
"""
import argparse, time, sys
from pathlib import Path
import win32com.client as win32
import pywintypes

def refresh_excel(file_path: Path, visible: bool=False, retries:int=3, delay:float=2.0):
    file_path = Path(file_path).resolve()
    last_err=None
    for attempt in range(1, retries+1):
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
            excel.Visible = visible
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(str(file_path), UpdateLinks=3, ReadOnly=False)
            # Refresh connections & pivots
            try:
                wb.RefreshAll()
                # give some time for background queries
                time.sleep(1.0)
            except Exception:
                pass
            wb.Save()
            wb.Close(SaveChanges=True)
            excel.Quit()
            print(f"REFRESH OK: {file_path}")
            return True
        except pywintypes.com_error as e:
            last_err = e
            print(f"[Attempt {attempt}/{retries}] Excel Open failed: {e}", file=sys.stderr)
            # try to close any zombie instance
            try:
                excel.Quit()
            except Exception:
                pass
            time.sleep(delay)
        except Exception as e:
            last_err = e
            print(f"[Attempt {attempt}/{retries}] Unexpected error: {e}", file=sys.stderr)
            try:
                excel.Quit()
            except Exception:
                pass
            time.sleep(delay)
    print(f"FAILED to refresh after {retries} attempts. Last error: {last_err}", file=sys.stderr)
    return False

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, help="Ścieżka do pliku XLSX")
    ap.add_argument("--visible", action="store_true")
    args = ap.parse_args()
    p = Path(args.file)
    if not p.exists():
        print(f"Plik nie istnieje: {p}", file=sys.stderr); sys.exit(1)
    ok = refresh_excel(p, visible=args.visible, retries=3, delay=2.0)
    sys.exit(0 if ok else 2)

if __name__ == "__main__":
    main()
