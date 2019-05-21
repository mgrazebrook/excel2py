"""
Demo of the prototype on how excel_to_py should work.

This module represents the hand-written code which uses it.
It also shows how you can override the generated calculations.
"""
from gen_demo import GeneratedProformaCalc


class CustomProformaCalc(GeneratedProformaCalc):
    # I want to provide my own version of an expression so I override the generated calc
    # In real code, this might represent data from the database.
    MATRIX = ((5, 32),
              (7, 8))


if __name__ == "__main__":
    calc = CustomProformaCalc()
    calc.calculate(Input_B3=6)

    # That's it!
    # both intermediate variables and results are available

    # Now to play with the results ...

    # From the Input tab - it would be nicer if it were named in Excel
    print("Input_B3:", calc.Input_B3)

    # Looking at intermediate values, perhaps to help debug it.
    print("BIGGER:", calc.BIGGER)
    print("MATRIX", repr(calc.MATRIX))

    # This is the value(s) from the Result tab.
    print("The Truth:", calc.TheTruth)
