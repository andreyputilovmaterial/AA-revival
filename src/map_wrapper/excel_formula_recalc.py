

import win32com.client as win32
from pathlib import Path

def refresh_and_save(fname_from, fname_to, config={}):
    """
    Open an Excel workbook, recalculate formulas, and save it again.
    
    Args:
        file_path (str): Full path to the Excel file.
    """
    print('{script_name}: re-calculating formulas...'.format(script_name=config['script_name']))
    print('{script_name}: launching Excel app do that...'.format(script_name=config['script_name']))
    if not Path(fname_from).is_file():
        raise FileNotFoundError(f"{fname_from} not found.")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # suppress prompts

    try:
        wb = excel.Workbooks.Open(fname_from)
        excel.CalculateFullRebuild()  # recalc all formulas
        wb.Save(fname_to)
        wb.Close()
        Path(fname_from).unlink()
        print('{script_name}: saved, and deleted the temporary file'.format(script_name=config['script_name']))
    finally:
        excel.Quit()
