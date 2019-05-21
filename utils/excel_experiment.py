"""
Junk file for experimenting with a live workbook
"""

import win32com.client as win32
import sys

xl = win32.gencache.EnsureDispatch('Excel.Application')
try:
    try:
        name = sys.argv[1]
        book = xl.Workbooks(name)
    except IndexError:
        print("excel_experiment <name of open spreadsheet>")
except win32.pywintypes.com_error:
    print(f"You must have '{name}' open in Excel first.")
