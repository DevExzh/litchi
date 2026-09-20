#!/usr/bin/env python3
"""Finish quality with explicit baseline-proven XLSB/PPTX test lint exceptions."""
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
    assert not (P/'quality-second-attempt.json').exists()
    initial = json.loads((P/'quality-initial.json').read_text())
    second = json.loads((P/'quality.json').read_text())
    assert [r['name'] for r in second] == ['fmt', 'tests', 'clippy-direct']
    assert [r['exit_code'] for r in second] == [0, 0, 101]
    for crate, candidate_log, message, locations in [
        ('xlsb', 'quality-clippy.log', 'error: used `expect()` on a `Result` value',
         ['comments/threaded/tests/mod.rs:271', 'comments/threaded/tests/mod.rs:276', 'comments/threaded/tests/mod.rs:384', 'shared_workbook/tests.rs:180']),
        ('pptx', 'quality-clippy-direct.log', 'error: called `.err().expect()` on a `Result` value',
         ['opened/tests.rs:464', 'opened/tests.rs:538', 'opened/tests.rs:557']),
    ]:
        baseline = json.loads((P/f'baseline-{crate}-clippy.json').read_text())
        assert baseline['exit_code'] == 101
        for name in [candidate_log, f'baseline-{crate}-clippy.log']:
            log = (P/name).read_text()
            assert log.count(message) == len(locations)
            assert all(location in log for location in locations)
    (P/'quality.json').rename(P/'quality-second-attempt.json')
    source = C.census()
    assert source == json.loads((P/'quality-source.json').read_text())
    assert source == json.loads((P/'source-candidate.json').read_text())
    common = ['--locked', '--target-dir', str(C.TARGET), '-j', '2']
    names = ['litchi-ooxml-common', 'litchi-drawingml', 'litchi-spreadsheet-drawing',
             'litchi-docx', 'litchi-xlsx']
    direct = [arg for name in names for arg in ['-p', name]]
    all_packages = [*direct, '-p', 'litchi-pptx', '-p', 'litchi-xlsb']
    jobs = [
        ('clippy-unaffected-targets', ['cargo', 'clippy', *direct, *common, '--all-features', '--all-targets', '--', '-D', 'warnings']),
        ('clippy-all-libs', ['cargo', 'clippy', *all_packages, *common, '--all-features', '--lib', '--', '-D', 'warnings']),
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
