#!/usr/bin/env python3
"""Retain instruction samples and disassembly for the frozen native probe."""
import hashlib
import json
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


build = json.loads((P / 'build.json').read_text())
profile = json.loads((P / 'profile-data.json').read_text())
binary = Path(build['binary'])
data = Path(profile['path'])
assert sha(binary) == build['binary_sha256']
assert sha(data) == profile['sha256']
out = P / 'instructions'
out.mkdir(exist_ok=True)
rows = []


def run(name, command):
    stdout = out / (name + '.stdout')
    stderr = out / (name + '.stderr')
    with stdout.open('w') as output, stderr.open('w') as error:
        result = subprocess.run(command, cwd=ROOT, stdout=output, stderr=error)
    rows.append(dict(name=name, command=command, exit_code=result.returncode,
                     stdout_sha256=sha(stdout), stderr_sha256=sha(stderr)))
    (P / 'instruction-runs.json').write_text(json.dumps(dict(
        binary_sha256=sha(binary), profile_data_sha256=sha(data), runs=rows
    ), indent=2) + '\n')
    assert result.returncode == 0, stderr.read_text()
    return stdout.read_text()


symbols = run('symbols', ['nm', '-S', '-C', str(binary)])
raw_symbols = run('raw-symbols', ['nm', '-S', str(binary)])
for name, symbol in [
    ('start', 'litchi_ooxml_common::mce::codec::start'),
    ('inherited-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Inherited>'),
    ('ctx-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Ctx>'),
]:
    matched = [line.split(maxsplit=3) for line in symbols.splitlines()
               if len(line.split(maxsplit=3)) == 4
               and line.split(maxsplit=3)[3] == symbol]
    assert len(matched) == 1, (symbol, matched)
    address = matched[0][0]
    raw = [line.split()[3] for line in raw_symbols.splitlines()
           if len(line.split()) == 4 and line.split()[0] == address]
    assert len(raw) == 1, (symbol, raw)
    run(name + '-assembly', ['objdump', '-d', '-C', '--no-show-raw-insn',
                            '--start-address=' + str(int(address, 16)),
                            '--stop-address=' + str(int(address, 16) + int(matched[0][1], 16)), str(binary)])
    run(name, ['perf', 'annotate', '--stdio', '--show-nr-samples',
               '--symbol', symbol, '-i', str(data)])
print('Frozen probe instruction evidence retained.')
