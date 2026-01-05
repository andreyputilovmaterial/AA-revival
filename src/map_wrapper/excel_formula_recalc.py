

import win32com.client
import time

from pathlib import Path


xl_constants = {
    'xlCalculationAutomatic': -4105,  # from enum XlCalculation
    'xlCalculationManual': -4135,  # from enum XlCalculation
    'xlCalculationSemiautomatic': 2,  # from enum XlCalculation
    'xlInterruption': {
        'xlAnyKey': 2,  # from enum XlCalculationInterruptKey
        'xlEscKey': 1,  # from enum XlCalculationInterruptKey
        'xlNoKey': 0,  # from enum XlCalculationInterruptKey
    },
    'xlCalculating': 1,  # from enum XlCalculationState
    'xlDone': 0,  # from enum XlCalculationState
    'xlPending': 2,  # from enum XlCalculationState
}


def refresh_and_save(fname_from, fname_to, config={}):
    """
    Open an Excel workbook, recalculate formulas, and save it again.
    
    Args:
        file_path (str): Full path to the Excel file.
    """
    print('{script_name}: re-calculating formulas...'.format(script_name=config['script_name']))
    print('Excel App: launching Excel app do that...')
    if not Path(fname_from).is_file():
        raise FileNotFoundError(f"{fname_from} not found.")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # suppress prompts

    try:
        print('Excel App: opening file...')
        wb = excel.Workbooks.Open(fname_from)
        print('Excel App: calling Calculate()...')
        # excel.CalculateFullRebuild()  # recalc all formulas
        for ws in wb.Sheets:
            ws.Calculate()
        excel.Calculation = xl_constants['xlCalculationAutomatic']
        while excel.CalculationState != xl_constants['xlDone']:
            print('Still calculating...')
            if excel.CalculationState==xl_constants['xlCalculating']:
                print('Calculating...', end='\r', flush=True)
            elif excel.CalculationState==xl_constants['xlPending']:
                print('Pending...', end='\r', flush=True)
            else:
                print('Excel status == constant {n}...'.format(n=excel.CalculationState), end='\r', flush=True)
            time.sleep(3)
        print('Excel App: saving updated file...')
        wb.SaveAs('{f}'.format(f=fname_to))
        wb.Close()
        Path(fname_from).unlink()
        print('{script_name}: saved, and deleted the temporary file'.format(script_name=config['script_name']))
    finally:
        print('Excel App: quit Excel instance')
        excel.Quit()
