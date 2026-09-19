#!/usr/bin/env python3
"""Independent integer references from exactly represented binary64 inputs.

Uses Python bigint math, never floating factorial ratios or production kernels.
Run with --check to compare the retained JSON byte for byte.
"""
import json
from fractions import Fraction
import math
from pathlib import Path
import random
import struct
import sys


def reference(name, values):
    integers = [int(x) for x in values]
    if name == "COMBIN":
        return math.comb(*integers)
    if name == "COMBINA":
        n, k = integers
        return math.comb(n + k - 1, k)
    if name == "FACT":
        return math.factorial(integers[0])
    if name == "FACTDOUBLE":
        return math.prod(range(integers[0], 0, -2))
    if name == "GCD":
        return math.gcd(*integers)
    if name == "LCM":
        return math.lcm(*integers)
    if name == "MULTINOMIAL":
        if any(x != math.floor(x) for x in values):
            numerator = math.factorial(math.floor(sum(map(Fraction, values))))
            denominator = math.prod(math.factorial(math.floor(x)) for x in values)
            assert numerator % denominator == 0
            return numerator // denominator
        total, result = 0, 1
        for value in integers:
            result *= math.comb(total + value, value)
            total += value
        return result
    if name in ("EVEN", "ODD"):
        magnitude = math.ceil(abs(values[0]))
        if magnitude % 2 != (name == "ODD"):
            magnitude += 1
        return -magnitude if values[0] < 0 else magnitude
    x, y = values if len(values) == 2 else (values[0], 0)
    return int(x == y) if name == "DELTA" else int(x >= y)


cases = []


def add(name, *values):
    cases.append((name, [float(x) for x in values]))


for n in [0, 1, 2, 20, 21, 50, 100, 169, 170, 171, 172]:
    add("FACT", n)
for n in [0, 1, 2, 3, 20, 100, *range(295, 306)]:
    add("FACTDOUBLE", n)
for n, k in [(5, 2), (171, 2), (1000, 3), (1000, 500), (1028, 514),
             (1029, 514), (2000, 1000), (2**53, 1), (2**53, 2),
             (2**53, 2**53 - 1), (2**54, 2**54 - 2),
             (1e308, 0), (1e308, 1), (1e308, 2), (1e308, 1e308)]:
    add("COMBIN", n, k)
for n, k in [(1, 0), (1, 1000), (2, 1000), (1000, 2), (1000, 500),
             (2**53, 1), (2**53, 2), (2**53, 3), (1e308, 1), (1e308, 2)]:
    add("COMBINA", n, k)
for values in [(0, 0), (0, 42), (48, 18, 30), (2**53, 2**53 - 1),
               (1e308, 1e308), (1e308, 2), (1e308, 3),
               (1e308, math.nextafter(1e308, 0)),
               (1e308, math.nextafter(1e308, 0), 0)]:
    add("GCD", *values)
    add("LCM", *values)
for values in [(0,), (1,), (2, 3, 4), (100, 100), (170, 1), (1000, 2),
               (2**53, 1, 1), (2**53, 2, 1), (1e308, 0), (1e308, 1),
               (1e308, 1, 1), tuple([1] * 40)]:
    add("MULTINOMIAL", *values)
for values in [(1.5, 1.5), (0.5, 0.5, 0.5, 0.5), (2.5, 1.5),
               (1.4, 0.6), (1.4, 0.6, 1), (0.1,) * 20,
               (0.25,) * 12, (0.75, 0.75, 0.75),
               (10.5, 10.5), (100.5, 70.5), (170, 0.5, 0.5),
               (0.9999999999999999, 1.0)]:
    add("MULTINOMIAL", *values)
for value in [-1e308, -2**53, -2.5, -1, -5e-324, 0, 5e-324,
              0.5, 1, 2, 2.5, 2**53 - 1, 2**53, 1e308]:
    for name in ["EVEN", "ODD"]:
        add(name, value)
for name in ["DELTA", "GESTEP"]:
    for values in [(0,), (-0.0,), (1,), (-1,), (1e308, 1e308),
                   (1, math.nextafter(1, 2)), (-1, -2)]:
        add(name, *values)

rng = random.Random(0xD15C4E7E)
for _ in range(32):
    n = rng.randrange(1, 1100)
    add("COMBIN", n, rng.randrange(n + 1))
    add("COMBINA", rng.randrange(1, 600), rng.randrange(600))
    add("FACT", rng.randrange(175))
    add("FACTDOUBLE", rng.randrange(310))
    values = [rng.randrange(1 << 48) for _ in range(rng.randrange(2, 6))]
    add("GCD", *values)
    add("LCM", *values)
    add("MULTINOMIAL", *(rng.randrange(40) for _ in range(rng.randrange(2, 7))))

rows = []
for name, values in cases:
    exact = reference(name, values)
    profile_error = None
    try:
        rounded = float(exact)
        bits = struct.pack(">d", rounded).hex() if math.isfinite(rounded) else None
    except OverflowError:
        bits = None
    rounded_bits = bits
    if name == "COMBINA" and values[0] < values[1]:
        profile_error = "Strict normative N >= M domain; optional extension not enabled"
        bits = None
    elif name == "ODD" and bits is not None and int(rounded) % 2 == 0:
        profile_error = "Exact odd result is not representable as an odd binary64 integer"
        bits = None
    rows.append({"function": name,
                 "formula": "=" + name + "(" + ";".join(map(repr, values)) + ")",
                 "input_binary64_hex": [x.hex() for x in values],
                 "exact_integer": str(exact), "rounded_reference_bits": rounded_bits,
                 "profile_error": profile_error, "expected_bits": bits})
document = {"oracle": "Python bigint math over exactly represented binary64 operands",
            "seed": "0xD15C4E7E", "observations": rows}
output = json.dumps(document, indent=2) + "\n"
path = Path(__file__).with_name("numeric-goldens.json")
if "--check" in sys.argv:
    if path.read_text() != output:
        raise SystemExit("retained numerical oracle differs")
    print(f"Verified {len(rows)} exact integer observations")
else:
    path.write_text(output)
    print(f"Wrote {len(rows)} exact integer observations")
