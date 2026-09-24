#!/usr/bin/env python3
"""Seeded binary64 aggregate cases checked with exact rational arithmetic."""
from fractions import Fraction
import json
import math
import random
import struct

rng = random.Random(0x0DF1407)
exponents = [-1074, -1073, -1050, -1022, -1000, -540, -538, -537,
             -500, -1, 0, 1, 500, 511, 512, 1000, 1023]


def number():
    mantissa = 1.0 + rng.getrandbits(52) / 2**52
    return math.ldexp(mantissa, rng.choice(exponents)) * rng.choice([-1.0, 1.0])


def array(values):
    return "{" + ";".join(repr(x) for x in values) + "}"


rows = []
names = ["SUM", "PRODUCT", "SUMSQ", "SUMPRODUCT", "SUMX2MY2", "SUMX2PY2", "SUMXMY2"]
for name in names:
    for index in range(48):
        left = [number() for _ in range(rng.randint(1, 6))]
        right = [number() for _ in left]
        if name == "SUM" and index % 3 == 0:
            left = [left[0], number(), -left[0]]
        if name == "SUMPRODUCT" and index % 3 == 0:
            left = [left[0], left[0], 1.5e-162]
            right = [right[0], -right[0], 1.5e-162]
        if name == "SUMX2MY2" and index % 3 == 0:
            right = left.copy()
            left = left + [1.0]
            right = right + [0.0]
        if name == "SUMXMY2" and index % 3 == 0:
            right = [math.nextafter(x, 0.0) for x in left]
        a, b = [Fraction(x) for x in left], [Fraction(x) for x in right]
        if name in ("SUM", "PRODUCT", "SUMSQ"):
            formula = "=" + name + "(" + ";".join(repr(x) for x in left) + ")"
            if name == "SUM":
                result = sum(a, Fraction())
            elif name == "PRODUCT":
                result = math.prod(a)
            else:
                result = sum((x*x for x in a), Fraction())
        else:
            formula = "=" + name + "(" + array(left) + ";" + array(right) + ")"
            pairs = zip(a, b, strict=True)
            if name == "SUMPRODUCT":
                result = sum((x*y for x, y in pairs), Fraction())
            elif name == "SUMX2MY2":
                result = sum((x*x-y*y for x, y in pairs), Fraction())
            elif name == "SUMX2PY2":
                result = sum((x*x+y*y for x, y in pairs), Fraction())
            else:
                result = sum(((x-y)**2 for x, y in pairs), Fraction())
        try:
            expected = float(result)
            bits = struct.pack(">d", expected).hex()
        except OverflowError:
            bits = None
        rows.append({"function": name, "case": index, "formula": formula,
                     "expected_bits": bits, "expected_error": "Number" if bits is None else None,
                     "max_ulps": 8 if name in ("PRODUCT", "SUMXMY2") else 0})
print(json.dumps({"seed": "0x0DF1407", "oracle": "fractions.Fraction",
                  "observations": rows}, indent=2))
