#!/usr/bin/env python3
"""Where the first changed byte of a published XLS edit lies (change 0748).

Prices the follow-up of forking the target digest from the source digest's
state at the first changed byte. Reads the census fixture list and the
directory of published outputs the census dumps with `XLS_DUMP_DIR`
(`{file name}-{operation}.bin`), skips ambiguous file names, no-ops and length
changes, and prints the distribution of `first changed offset / length` per
operation.

usage: first_change.py REPO_ROOT FIXTURES_TXT DUMP_DIR
"""
import os
import statistics
import sys

ROOT, FIXTURES, DUMP = sys.argv[1:4]
OPERATIONS = ('number-plan', 'number-source-backed', 'number-generic', 'string-generic', 'noop-generic')


def main():
    paths = [line.strip() for line in open(FIXTURES) if line.strip()]
    by_name = {}
    for path in paths:
        by_name.setdefault(path.rsplit('/', 1)[-1], []).append(path)
    fractions = {}
    for name in sorted(os.listdir(DUMP)):
        stem = operation = None
        for candidate in OPERATIONS:
            suffix = f'-{candidate}.bin'
            if name.endswith(suffix):
                stem, operation = name[:-len(suffix)], candidate
        if stem is None or len(by_name.get(stem, [])) != 1:
            continue
        source = open(os.path.join(ROOT, by_name[stem][0]), 'rb').read()
        output = open(os.path.join(DUMP, name), 'rb').read()
        if len(source) != len(output) or source == output:
            continue
        first = next(index for index in range(len(source)) if source[index] != output[index])
        fractions.setdefault(operation, []).append((first / len(source), stem, first, len(source)))
    for operation, values in sorted(fractions.items()):
        ratios = sorted(value[0] for value in values)
        print(f'{operation}: n={len(ratios)} median {statistics.median(ratios):.3f} '
              f'p25 {ratios[len(ratios) // 4]:.3f} p75 {ratios[3 * len(ratios) // 4]:.3f}')
        for ratio, stem, first, length in sorted(values, key=lambda value: value[1]):
            if stem == '54016.xls':
                print(f'  54016.xls first changed byte {first} of {length} ({ratio:.3f})')


if __name__ == '__main__':
    main()
