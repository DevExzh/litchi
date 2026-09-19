#!/usr/bin/env python3
"""Reproduce extreme goldens from exact binary64 inputs; no library dependency."""
import json
import math
import struct
import sys
from fractions import Fraction

import mpmath as mp

assert mp.__version__ == "1.3.0"
mp.mp.dps = 600
cases = [
    ("SQRTPI", [sys.float_info.max], lambda x: mp.sqrt(x * mp.pi)),
    ("SQRTPI", [5e-324], lambda x: mp.sqrt(x * mp.pi)),
    ("LN", [5e-324], mp.log),
    ("LN", [math.nextafter(1.0, 2.0)], mp.log),
    ("LOG10", [5e-324], mp.log10),
    ("LOG", [5e-324, math.nextafter(1.0, 2.0)], lambda x, b: mp.log(x) / mp.log(b)),
    ("LOG", [math.nextafter(1.0, 2.0), math.nextafter(1.0, 0.0)], lambda x, b: mp.log(x) / mp.log(b)),
    ("EXP", [-745.0], mp.exp),
    ("EXP", [709.0], mp.exp),
]
observations = []
for name, inputs, function in cases:
    result = function(*(mp.mpf(value) for value in inputs))
    rounded = float(result)
    observations.append({
        "function": name, "input_decimal": [repr(x) for x in inputs],
        "input_binary64_hex": [x.hex() for x in inputs],
        "reference_decimal": mp.nstr(result, 45),
        "rounded_binary64_bits": struct.pack(">d", rounded).hex(),
    })
for a, b in [(1e308, 3.0), (-1e308, 3.0), (1e308, -3.0),
             (sys.float_info.max, 1e-300), (26.0 ** 15, 77.0),
             (-5e-324, sys.float_info.max)]:
    x, y = Fraction(a), Fraction(b)
    result = x - (x // y) * y
    observations.append({
        "function": "MOD", "input_decimal": [repr(a), repr(b)],
        "input_binary64_hex": [a.hex(), b.hex()],
        "exact_rational_numerator": str(result.numerator),
        "exact_rational_denominator": str(result.denominator),
        "rounded_binary64_bits": struct.pack(">d", float(result)).hex(),
    })
print(json.dumps({"oracle": "mpmath 1.3.0 and fractions.Fraction",
                  "decimal_precision": 600, "observations": observations}, indent=2))
