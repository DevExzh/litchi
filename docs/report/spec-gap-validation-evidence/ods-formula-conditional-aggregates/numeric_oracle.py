#!/usr/bin/env python3
"""Independent selection masks and exact represented-binary64 reductions.

Standard-library only. No Rust evaluator output is consulted. Numeric criteria
use integer fixture columns, while selected values cover binary64 extremes.
"""
from fractions import Fraction
from pathlib import Path
import json
import math
import random
import struct
import sys

HERE = Path(__file__).resolve().parent

def bits(value):
    return struct.pack('>d', value).hex()

def document():
    rng = random.Random(20260919)
    maxval = sys.float_info.max
    tiny = math.ulp(0.0)
    vectors = [
        [maxval, 3.0, -maxval], [tiny, maxval, -maxval],
        [maxval, maxval], [tiny, 0.0], [tiny, tiny, tiny, 0.0],
        [1e16, 1.0, -1e16], [-tiny, 0.0], [0.0, -0.0],
    ]
    for _ in range(40):
        vectors.append([math.ldexp(rng.choice([-1, 1]) * rng.random(), rng.randint(-1000, 1000))
                        for _ in range(rng.randint(2, 45))])
    rows = []
    for index, values in enumerate(vectors):
        # Every fourth vector has no matches; first eight retain all extremes.
        n = len(values)
        a = [1 if index < 8 else rng.randrange(-2, 4) for _ in values]
        b = [1 if index < 8 else rng.randrange(0, 3) for _ in values]
        threshold = 99 if index >= 8 and index % 4 == 0 else 1
        single = [i for i in range(n) if a[i] >= threshold]
        plural = [i for i in single if b[i] == 1]
        cells = [[bits(float(x)), bits(float(y)), bits(z)] for x, y, z in zip(a, b, values)]
        for function in ('SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'AVERAGEIFS'):
            indices = plural if function.endswith('IFS') else single
            args = f'[.A1:.A{n}];">={threshold}"'
            if function.endswith('IFS'):
                args += f';[.B1:.B{n}];1'
                if not function.startswith('COUNT'):
                    args = f'[.C1:.C{n}];' + args
            elif not function.startswith('COUNT'):
                # Destination shape intentionally a single anchor, not n rows.
                args += ';[.C1]'
            row = {'case': index, 'function': function, 'formula': f'={function}({args})',
                   'cells_bits': cells, 'selected_indices': indices}
            if function.startswith('COUNT'):
                exact = Fraction(len(indices))
            else:
                exact = sum((Fraction(values[i]) for i in indices), Fraction())
                if function.startswith('AVERAGE'):
                    if not indices:
                        row['error'] = 'DivisionByZero'
                        rows.append(row)
                        continue
                    exact /= len(indices)
            row.update(numerator=str(exact.numerator), denominator=str(exact.denominator))
            try:
                rounded = float(exact)
                row['expected_bits'] = bits(rounded)
            except OverflowError:
                row['error'] = 'Number'
            rows.append(row)
    return {'oracle': 'Python Fraction from exact binary64 operands; independent integer criteria masks',
            'seed': 20260919, 'observations': rows}

def main():
    encoded = json.dumps(document(), indent=2) + '\n'
    path = HERE / 'numeric-goldens.json'
    if '--check' in sys.argv:
        assert path.read_text() == encoded, 'retained oracle differs'
        print(json.dumps({'observations': len(document()['observations']), 'verified': True}))
    else:
        path.write_text(encoded)

if __name__ == '__main__':
    main()
