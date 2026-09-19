#!/usr/bin/env python3
"""Bind native ownership symbols and their address-bounded disassembly."""
import hashlib
import json
import subprocess
import sys
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
phase = sys.argv[1]
assert phase in ['baseline', 'candidate']
build = next(row for row in json.loads((P / ('builds-' + phase + '.json')).read_text())
             if row['label'] == 'native')
binary = Path(build['binary'])
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
assert sha(binary) == build['binary_sha256']
out = P / 'ownership-assembly'
out.mkdir(exist_ok=True)
commands = []


def run(name, command):
    result = subprocess.run(command, capture_output=True, text=True)
    stdout = out / (phase + '-' + name + '.stdout')
    stderr = out / (phase + '-' + name + '.stderr')
    stdout.write_text(result.stdout)
    stderr.write_text(result.stderr)
    commands.append(dict(name=name, command=command, exit_code=result.returncode,
                         stdout_sha256=sha(stdout), stderr_sha256=sha(stderr)))
    assert result.returncode == 0, result.stderr
    return result.stdout


table = run('symbols', ['nm', '-S', '-C', str(binary)])
symbols = []
for name, symbol in [
    ('start', 'litchi_ooxml_common::mce::codec::start'),
    ('inherited-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Inherited>'),
    ('ctx-drop', 'core::ptr::drop_in_place<litchi_ooxml_common::mce::codec::Ctx>'),
]:
    matches = [line.split(maxsplit=3) for line in table.splitlines()
               if len(line.split(maxsplit=3)) == 4 and line.split(maxsplit=3)[3] == symbol]
    assert len(matches) <= 1, symbol
    if not matches:
        assert name != 'start'
        symbols.append(dict(name=name, symbol=symbol, present=False))
        continue
    address, size = [int(value, 16) for value in matches[0][:2]]
    output = run(name, ['objdump', '-d', '-C', '--no-show-raw-insn',
                        '--start-address=' + str(address),
                        '--stop-address=' + str(address + size), str(binary)])
    assert '<' + symbol + '>:' in output, symbol
    symbols.append(dict(name=name, symbol=symbol, present=True,
                        address=address, size=size))
(out / (phase + '.json')).write_text(json.dumps(dict(
    phase=phase, binary=str(binary), binary_sha256=sha(binary),
    symbols=symbols, commands=commands,
), indent=2) + '\n')
print(phase, 'ownership symbol evidence retained')
