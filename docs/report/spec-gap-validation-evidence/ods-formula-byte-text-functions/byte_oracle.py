#!/usr/bin/env python3
"""Independent UTF-8 boundary oracle for the seven byte-position functions."""
import argparse
import hashlib
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
FUNCTIONS = ('FINDB', 'LEFTB', 'LENB', 'MIDB', 'REPLACEB', 'RIGHTB', 'SEARCHB')


def boundaries(text):
    result = [0]
    for character in text:
        result.append(result[-1] + len(character.encode('utf-8')))
    return result


def start_index(offsets, position):
    """Map a one-based byte position backward to its containing scalar."""
    return max(index for index, offset in enumerate(offsets) if offset <= position - 1)


def span_end(offsets, start, length):
    return max(index for index in range(start, len(offsets))
               if offsets[index] - offsets[start] <= length)


def result(kind, value):
    return {'kind': kind, 'value': value}


def evaluate(name, args):
    if name in ('FINDB', 'SEARCHB'):
        needle, text, *optional = args
        position = optional[0] if optional else 1
        offsets = boundaries(text)
        if position < 1 or position > offsets[-1] + 1:
            return result('Error', 'Value')
        start = start_index(offsets, position)
        if needle == '':
            return result('Number', offsets[start] + 1)
        # Enumerate complete original-scalar spans. This deliberately differs
        # from the production streaming KMP algorithm and does not use its tables.
        folded = name == 'SEARCHB'
        wanted = needle.casefold() if folded else needle
        for first in range(start, len(text)):
            for end in range(first + 1, len(text) + 1):
                candidate = text[first:end]
                if (candidate.casefold() if folded else candidate) == wanted:
                    return result('Number', offsets[first] + 1)
        return result('Error', 'Value')
    text = args[0]
    offsets = boundaries(text)
    if name == 'LENB':
        return result('Number', offsets[-1])
    if name in ('LEFTB', 'RIGHTB'):
        length = args[1] if len(args) > 1 else 1
        if length < 0:
            return result('Error', 'Value')
        if name == 'LEFTB':
            value = text[:span_end(offsets, 0, length)]
        else:
            first = min(index for index, offset in enumerate(offsets)
                        if offsets[-1] - offset <= length)
            value = text[first:]
        return result('Text', value)
    position, length = args[1:3]
    if position < 1 or length < 0:
        return result('Error', 'Value')
    first = start_index(offsets, min(position, offsets[-1] + 1))
    end = span_end(offsets, first, length)
    if name == 'MIDB':
        return result('Text', text[first:end])
    assert name == 'REPLACEB'
    return result('Text', text[:first] + args[3] + text[end:])


def formula(name, args):
    def encode(value):
        return '"' + value.replace('"', '""') + '"' if isinstance(value, str) else str(value)
    return '=' + name + '(' + ';'.join(map(encode, args)) + ')'


def document():
    cases = []
    def add(name, *args):
        cases.append({'function': name, 'args': list(args), 'formula': formula(name, args),
                      'expected': evaluate(name, args)})
    # Every offset and short length for representative 1/2/3/4-byte scalars.
    for text in ('', 'abc', 'é', '界', '🙂', 'Aé界🙂', 'e\u0301x'):
        size = len(text.encode('utf-8'))
        add('LENB', text)
        for name in ('LEFTB', 'RIGHTB'):
            add(name, text)
            for length in range(-1, size + 3):
                add(name, text, length)
        for start in range(0, size + 3):
            for length in range(-1, size + 2):
                add('MIDB', text, start, length)
                add('REPLACEB', text, start, length, 'Ω')
        for name in ('FINDB', 'SEARCHB'):
            for needle in ('', 'x', 'é', '界', '🙂'):
                for start in range(0, size + 3):
                    add(name, needle, text, start)
    for needle, text in (('ss', 'aß'), ('s', 'ß'), ('ß', 'ss'), ('STRASSE', 'Straße'),
                         ('i', 'İ'), ('i\u0307', 'İ'), ('σ', 'ς'), ('é', 'e\u0301')):
        for start in range(1, len(text.encode('utf-8')) + 2):
            add('SEARCHB', needle, text, start)
    return {'contract_sha256': hashlib.sha256((HERE / 'contract.md').read_bytes()).hexdigest(),
            'profile': 'UTF-8; integer parameters; complete original-scalar spans',
            'casefold_scope': 'Python default full fold on selected stable characters; no claim of full Unicode17 coverage',
            'observations': cases}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--write', action='store_true')
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    data = document()
    payload = (json.dumps(data, ensure_ascii=False, indent=2) + '\n').encode()
    path = HERE / 'byte-goldens.json'
    if args.write:
        path.write_bytes(payload)
    if args.check:
        assert path.read_bytes() == payload, 'oracle bytes differ'
    print(json.dumps({'functions': len({r['function'] for r in data['observations']}),
                      'observations': len(data['observations']), 'verified': args.check}))


if __name__ == '__main__':
    main()
