from excel2py.base_proforma_calc import BaseProformaCalc
import unittest


class Sub(BaseProformaCalc):
    inputs = {'x'}


class TestBaseProformaCalc(unittest.TestCase):
    def test_good(self):
        s = Sub()
        s.calculate(x=4)

    def test_missing(self):
        s = Sub()
        with self.assertRaises(TypeError):
            s.calculate()

    def test_extra(self):
        s = Sub()
        with self.assertRaises(TypeError):
            s.calculate(x=5, y=3, z=4)


if __name__ == "__main__":
    unittest.main()
