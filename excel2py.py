"""
excel2py - turn a spreadsheet open in Excel into a Python program

It is controlled by a config file, traditionally a yaml file with
the same name as the Excel file. This tells it which sheets are
inputs, outputs and calculations.

By Michael Grazebrook of Joined Up Finance Ltd
for Willis Towers Watson, 21 Aug 2018
"""

from excel2py.config import config
from excel2py.excel_to_py import ExcelToPy


def main():
    cfg = config(__doc__)
    app = ExcelToPy(cfg)
    app.generate()


if __name__ == "__main__":
    main()
