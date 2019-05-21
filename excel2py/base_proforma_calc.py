"""
Abstract base class for excel2py's generated code.

By Michael Grazebrook of Joined Up Finance Ltd
"""


class BaseProformaCalc:
    """
    Abstract base class for proforma calculations.
    """
    inputs = set()

    def check_inputs(self, **args):
        keys = set(args.keys())
        missing = self.inputs - keys
        if missing:
            raise TypeError(f"calculate() missing {len(missing)} required arguments {','.join(missing)}")
        extra = keys - self.inputs
        if extra:
            raise TypeError(f"calculate() got {len(extra)} extra arguments {','.join(extra)}")

    def calculate(self, **args):
        self.check_inputs(**args)
        self.__dict__.update(args)
        # TODO: Return just the outputs, though this class has all outputs as attributes anyway

