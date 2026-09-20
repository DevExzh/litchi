#!/usr/bin/env python3
"""Finish quality after proving preexisting XLSB test-only Clippy debt."""
import importlib.util
import json
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('custody0713followup', P/'custody.py')
C = importlib.util.module_from_spec(spec)
spec.loader.exec_module(C)

def main():
    assert not (P/'quality-initial.json').exists()
    initial = json.loads((P/'quality.json').read_text())
    assert [r['name'] for r in initial] == ['fmt', 'tests', 'clippy']
    assert [r['exit_code'] for r in initial] == [0, 0, 101]
    baseline = json.loads((P/'baseline-xlsb-clippy.json').read_text())
    assert baseline['exit_code'] == 101
    for name in ['quality-clippy.log', 'baseline-xlsb-clippy.log']:
        log = (P/name).read_text()
        assert log.count('error: used `expect()` on a `Result` value') == 4
        for location in ['comments/threaded/tests/mod.rs:271',
                         'comments/threaded/tests/mod.rs:276',
                         'comments/threaded/tests/mod.rs:384',
                         'shared_workbook/tests.rs:180']:
            assert location in log
    (P/'quality.json').rename(P/'quality-initial.json')
    source = C.census()
    assert source == json.loads((P/'quality-source.json').read_text())
    assert source == json.loads((P/'source-candidate.json').read_text())
    common = ['--locked', '--target-dir', str(C.TARGET), '-j', '2']
    names = ['litchi-ooxml-common', 'litchi-drawingml', 'litchi-spreadsheet-drawing',
             'litchi-docx', 'litchi-xlsx', 'litchi-pptx']
    direct = [arg for name in names for arg in ['-p', name]]
    all_packages = [*direct, '-p', 'litchi-xlsb']
    jobs = [
        ('clippy-direct', ['cargo', 'clippy', *direct, *common, '--all-features', '--all-targets', '--', '-D', 'warnings']),
        ('clippy-xlsb-lib', ['cargo', 'clippy', '-p', 'litchi-xlsb', *common, '--all-features', '--lib', '--', '-D', 'warnings']),
        ('doctests', ['cargo', 'test', *all_packages, *common, '--all-features', '--doc']),
        ('rustdoc', ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', *all_packages, *common, '--all-features', '--no-deps']),
    ]
    rows = initial[:2]
    for name, cmd in jobs:
        log = P/('quality-'+name+'.log')
        assert not log.exists()
        start = time.monotonic()
        with log.open('w') as stream:
            result = subprocess.run(cmd, cwd=C.ROOT, stdout=stream, stderr=subprocess.STDOUT)
        assert C.census() == source
        rows.append(dict(name=name, command=cmd, exit_code=result.returncode,
                         seconds=time.monotonic()-start, log=log.name,
                         log_sha256=C.sha(log), source_manifest_sha256=C.sha(P/'quality-source.json')))
        (P/'quality.json').write_text(json.dumps(rows, indent=2)+'\n')
        print(name, result.returncode, flush=True)
        assert result.returncode == 0

if __name__ == '__main__':
    main()
