#!/usr/bin/env python3
"""Independent typed reference reductions; exact binary64 rational averages."""
from fractions import Fraction
from pathlib import Path
import json
import math
import random
import struct
import sys

HERE = Path(__file__).resolve().parent
FUNCTIONS = ('COUNT', 'COUNTA', 'COUNTBLANK', 'AVERAGE', 'AVERAGEA', 'MIN', 'MAX', 'MINA', 'MAXA')


def bits(value):
    return struct.pack('>d', value).hex()


def number(value):
    return {'kind': 'Number', 'bits': bits(value)}


def document():
    rng = random.Random(20260919)
    maximum, tiny = sys.float_info.max, math.ulp(0.0)
    vectors = [[maximum, 3., -maximum], [tiny, maximum, -maximum],
               [maximum, maximum], [tiny, 0.], [tiny, tiny, tiny, 0.],
               [1e16, 1., -1e16], [-tiny, 0.], [0., -0.]]
    for _ in range(40):
        vectors.append([math.ldexp(rng.choice([-1, 1]) * rng.random(), rng.randint(-1000, 1000))
                        for _ in range(rng.randint(2, 45))])
    datasets = [[number(x) for x in row] for row in vectors]
    empty = {'kind': 'Empty'}
    text = lambda value: {'kind': 'Text', 'value': value}
    logical = lambda value: {'kind': 'Logical', 'value': value}
    error = lambda value: {'kind': 'Error', 'value': value}
    datasets += [
        [empty], [text('')], [text('3')], [text('word')],
        [logical(True)], [logical(False)], [empty, text(''), text('3'), text('word')],
        [number(5.), logical(True), logical(False), text('3'), empty],
        [number(-5.), text(''), empty], [number(5.), text(''), empty],
        [error('NotAvailable')], [number(1.), error('DivisionByZero'), error('NotAvailable')],
        [empty, text(''), logical(False), error('Number'), number(0.)],
        [number(maximum), number(maximum), text('word'), empty],
        [number(-tiny), logical(False)],
        [number(tiny), empty, text('word'), logical(True), number(-tiny)],
    ]
    observations = []
    for case, cells in enumerate(datasets):
        for function in FUNCTIONS:
            row = {'case': case, 'function': function,
                   'formula': f'={function}([.A1:.A{len(cells)}])', 'cells': cells}
            if function == 'COUNT':
                exact = Fraction(sum(c['kind'] == 'Number' for c in cells))
            elif function == 'COUNTA':
                exact = Fraction(sum(c['kind'] != 'Empty' for c in cells))
            elif function == 'COUNTBLANK':
                exact = Fraction(sum(c['kind'] == 'Empty' or c == text('') for c in cells))
            else:
                errors = [c['value'] for c in cells if c['kind'] == 'Error']
                if errors:
                    row['error'] = errors[0]
                    observations.append(row)
                    continue
                values = []
                for cell in cells:
                    if cell['kind'] == 'Number':
                        values.append(Fraction(struct.unpack('>d', bytes.fromhex(cell['bits']))[0]))
                    elif function.endswith('A') and cell['kind'] == 'Logical':
                        values.append(Fraction(int(cell['value'])))
                    elif function.endswith('A') and cell['kind'] == 'Text':
                        values.append(Fraction())
                if function.startswith('AVERAGE'):
                    if not values:
                        row['error'] = 'DivisionByZero'
                        observations.append(row)
                        continue
                    exact = sum(values, Fraction()) / len(values)
                elif function.startswith('MIN'):
                    exact = min(values, default=Fraction())
                else:
                    exact = max(values, default=Fraction())
            row.update(numerator=str(exact.numerator), denominator=str(exact.denominator),
                       expected_bits=bits(float(exact)))
            observations.append(row)
    assert len(observations) == 576
    return {'oracle': 'Python Fraction from exact binary64 operands; independent typed reference admission',
            'seed': 20260919, 'countblank_empty_text_is_blank': True, 'observations': observations}


def main():
    result = document()
    encoded = json.dumps(result, indent=2) + '\n'
    path = HERE / 'numeric-goldens.json'
    if '--check' in sys.argv:
        assert path.read_text() == encoded, 'retained oracle differs'
        print(json.dumps({'observations': len(result['observations']), 'verified': True}))
    else:
        path.write_text(encoded)


if __name__ == '__main__':
    main()
