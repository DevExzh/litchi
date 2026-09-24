#!/usr/bin/env python3
"""Exact-rational aggregate goldens from the represented binary64 operands.

This evidence tool uses only Python's standard library. Product goldens describe
the mathematical reference; a scaled finite-mantissa implementation may round
during multiplication, so its documented comparison tolerance still applies.
"""
from fractions import Fraction
import json
import math
import struct
import sys

MAX = sys.float_info.max
cases = [
    ("SUM", "=SUM(1e16;1;-1e16)", [[1e16, 1.0, -1e16]]),
    ("SUM", "=SUM(5e-324;1e308;-1e308)", [[5e-324, 1e308, -1e308]]),
    ("SUM", "=SUM(1.7976931348623157e308;1.7976931348623157e308;-1.7976931348623157e308)", [[MAX, MAX, -MAX]]),
    ("PRODUCT", "=PRODUCT(1e308;1e308;1e-308;1e-308)", [[1e308, 1e308, 1e-308, 1e-308]]),
    ("PRODUCT", "=PRODUCT(1e-308;1e-308;1e308;1e308)", [[1e-308, 1e-308, 1e308, 1e308]]),
    ("SUMSQ", "=SUMSQ(1.5e-162;1.5e-162)", [[1.5e-162, 1.5e-162]]),
    ("SUMSQ", "=SUMSQ(1e-162;1e-162;1e-162)", [[1e-162] * 3]),
    ("SUMX2MY2", "=SUMX2MY2({1.7976931348623157e308;1};{1.7976931348623157e308;0})", [[MAX, 1.0], [MAX, 0.0]]),
    ("SUMX2PY2", "=SUMX2PY2({1e-162;1e-162};{1e-162;1e-162})", [[1e-162] * 2, [1e-162] * 2]),
    ("SUMXMY2", f"=SUMXMY2({{1e154}};{{{math.nextafter(1e154, 0.0)!r}}})", [[1e154], [math.nextafter(1e154, 0.0)]]),
    ("SUMPRODUCT", "=SUMPRODUCT({1.7976931348623157e308;1.7976931348623157e308};{2;-2})", [[MAX, MAX], [2.0, -2.0]]),
    ("SUMPRODUCT", "=SUMPRODUCT({1.7976931348623157e308;1.7976931348623157e308;1};{1.7976931348623157e308;-1.7976931348623157e308;1})", [[MAX, MAX, 1.0], [MAX, -MAX, 1.0]]),
    ("SUMPRODUCT", "=SUMPRODUCT({1.5e-162;1.5e-162};{1.5e-162;1.5e-162})", [[1.5e-162] * 2, [1.5e-162] * 2]),
    ("SUMPRODUCT", "=SUMPRODUCT({1.7976931348623157e308};{1.7976931348623157e308};{1e-309})", [[MAX], [MAX], [1e-309]]),
    ("SUMPRODUCT", "=SUMPRODUCT({1e308;1e308;1};{1e308;-1e308;1};{1e308;1e308;1})", [[1e308, 1e308, 1.0], [1e308, -1e308, 1.0], [1e308, 1e308, 1.0]]),
    ("SUMPRODUCT", f"=SUMPRODUCT({{1;{2.0**-100!r};{2.0**-200!r};-1;{-2.0**-100!r}}};{{1;1;1;1;1}};{{1;1;1;1;1}})", [[1.0, 2.0**-100, 2.0**-200, -1.0, -2.0**-100], [1.0]*5, [1.0]*5]),
    ("SUMPRODUCT", f"=SUMPRODUCT({{{2.0**-1000!r};{-math.nextafter(2.0**-1000, 0.0)!r}}};{{1;1}};{{1;1}})", [[2.0**-1000, -math.nextafter(2.0**-1000, 0.0)], [1.0]*2, [1.0]*2]),
]


def reference(name, operands):
    a = [[Fraction(x) for x in row] for row in operands]
    if name == "SUM":
        return sum(a[0], Fraction())
    if name == "PRODUCT":
        return math.prod(a[0])
    if name == "SUMSQ":
        return sum((x * x for x in a[0]), Fraction())
    if name == "SUMPRODUCT":
        return sum((math.prod(column) for column in zip(*a, strict=True)), Fraction())
    pairs = zip(a[0], a[1], strict=True)
    if name == "SUMX2MY2":
        return sum((x * x - y * y for x, y in pairs), Fraction())
    if name == "SUMX2PY2":
        return sum((x * x + y * y for x, y in pairs), Fraction())
    return sum(((x - y) ** 2 for x, y in pairs), Fraction())


rows = []
for name, formula, operands in cases:
    result = reference(name, operands)
    rows.append({
        "function": name, "formula": formula,
        "input_binary64_hex": [[x.hex() for x in group] for group in operands],
        "numerator": str(result.numerator), "denominator": str(result.denominator),
        "rounded_binary64_bits": struct.pack(">d", float(result)).hex(),
    })
print(json.dumps({"oracle": "fractions.Fraction from exact binary64 operands",
                  "observations": rows}, indent=2))
