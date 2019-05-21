"""
This hand-crafted class is a prototype of how I'd like the generated output to work.
Output fro demo.xlsx should look just like this.
"""
from excel_functions import *
from base_proforma_calc import BaseProformaCalc


class GeneratedProformaCalc(BaseProformaCalc):  # though you'd probably give it a name
    inputs = {'Input_B3'}
    outputs = {'TheTruth'}

    # A cell given the name A
    # Can be overridden in the customisation sub-class.
    A = 2

    # A named range containing values
    MATRIX = ((1, 2),
              (3, 4))

    # I haven't yet decided how to implement a named range containing formulas
    # It should be possible to make a function out of it if all cells are simply related

    # A typical named cell with a formula in it
    @property
    def BIGGER(self):
        return MAX(self.A, 4)

    # This is what happens if you forget to name a cell
    @property
    def Calc_B9(self):
        return self.BIGGER + self.Input_B3

    # Outputs are just properties. ALL OUTPUTS MUST BE NAMED CELLS
    @property
    def TheTruth(self):
        return VLOOKUP(self.Input_B3, self.MATRIX, 2) + self.Calc_B9
