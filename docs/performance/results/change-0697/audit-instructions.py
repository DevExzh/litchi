#!/usr/bin/env python3
"""Independently check retained instruction commands, identities and bytes."""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


record = json.loads((P / 'instruction-runs.json').read_text())
build = json.loads((P / 'build.json').read_text())
profile = json.loads((P / 'profile-data.json').read_text())
assert record['binary_sha256'] == build['binary_sha256']
assert record['profile_data_sha256'] == profile['sha256']
for path, digest, key in [(Path(build['binary']), build['binary_sha256'], 'binary_sha256'),
                          (Path(profile['path']), profile['sha256'], 'profile_data_sha256')]:
    if path.is_file():
        assert sha(path) == digest
    else:
        cleanup = json.loads((P / 'cleanup.json').read_text())
        assert cleanup[key] == digest
expected = [('symbols', ['nm', '-S', '-C', build['binary']]),
            ('raw-symbols', ['nm', '-S', build['binary']])]
symbols = (P / 'instructions/symbols.stdout').read_text().splitlines()
raw_symbols = (P / 'instructions/raw-symbols.stdout').read_text().splitlines()
for name, symbol in [
    ('start', 'litchi_ooxml_common::mce::codec::start'),
    ('inherited-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Inherited>'),
    ('ctx-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Ctx>'),
]:
    matches = [line.split(maxsplit=3) for line in symbols
               if len(line.split(maxsplit=3)) == 4 and line.split(maxsplit=3)[3] == symbol]
    assert len(matches) == 1
    names = [line.split()[3] for line in raw_symbols
             if len(line.split()) == 4 and line.split()[0] == matches[0][0]]
    assert len(names) == 1
    expected += [(name + '-assembly', ['objdump', '-d', '-C', '--no-show-raw-insn',
                                      '--start-address=' + str(int(matches[0][0], 16)),
                                      '--stop-address=' + str(int(matches[0][0], 16) + int(matches[0][1], 16)), build['binary']]),
                 (name, ['perf', 'annotate', '--stdio', '--show-nr-samples',
                         '--symbol', symbol, '-i', profile['path']])]
    assembly = (P / 'instructions' / (name + '-assembly.stdout')).read_text()
    assert '<' + symbol + '>:' in assembly, name
    assert 'lock ' in assembly, name
    annotation = (P / 'instructions' / (name + '.stdout')).read_text()
    assert symbol in annotation and len(annotation.splitlines()) > 10, name
assert len(record['runs']) == len(expected)
for row, (name, command) in zip(record['runs'], expected):
    assert row['name'] == name and row['command'] == command
    assert row['exit_code'] == 0
    for extension in ['stdout', 'stderr']:
        assert sha(P / 'instructions' / (name + '.' + extension)) == row[extension + '_sha256']
before = sha(P / 'instruction-summary.json')
subprocess.run([sys.executable, str(P / 'summarize-instructions.py')], check=True,
               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
assert sha(P / 'instruction-summary.json') == before
print('PASS: frozen instruction symbols, commands, assembly, annotations and cleanup bindings')
