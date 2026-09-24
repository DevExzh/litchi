#!/usr/bin/env python3
"""Reproduce test goldens with mpmath 1.3.0 at 600 decimal places.

Run this optional evidence tool with that version installed; mpmath is not a
Rust/runtime dependency. Inputs are converted from exact binary64 values.
"""
import json
import struct

import mpmath as mp

assert mp.__version__ == "1.3.0"
mp.mp.dps = 600
cases = [
    ("ACOTH", 1e308, lambda x: mp.log((x + 1) / (x - 1)) / 2),
    ("ACOTH", -1e308, lambda x: mp.log((x + 1) / (x - 1)) / 2),
    ("ACOSH", 1e308, mp.acosh),
    ("ASINH", 1e308, mp.asinh),
    ("ASINH", -1e308, mp.asinh),
    ("SECH", 710.0, lambda x: 1 / mp.cosh(x)),
    ("CSCH", 710.0, lambda x: 1 / mp.sinh(x)),
    ("SECH", 720.0, lambda x: 1 / mp.cosh(x)),
    ("CSCH", 720.0, lambda x: 1 / mp.sinh(x)),
    ("SECH", 745.5, lambda x: 1 / mp.cosh(x)),
    ("CSCH", 745.5, lambda x: 1 / mp.sinh(x)),
    ("SECH", 745.8, lambda x: 1 / mp.cosh(x)),
    ("CSCH", 745.8, lambda x: 1 / mp.sinh(x)),
]
observations = []
for name, value, function in cases:
    result = function(mp.mpf(value))
    rounded = float(result)
    observations.append({
        "function": name, "input_binary64_hex": value.hex(),
        "reference_decimal": mp.nstr(result, 45),
        "rounded_binary64_hex": rounded.hex(),
        "rounded_binary64_bits": struct.pack(">d", rounded).hex(),
    })
print(json.dumps({"oracle": "mpmath 1.3.0", "decimal_precision": 600,
                  "observations": observations}, indent=2))
