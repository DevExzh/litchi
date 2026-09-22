#!/usr/bin/env python3
"""Freeze inputs, then run the prospective serial unchanged-owner matrix."""
import subprocess
import sys
import time
from contract import P, ROOT, OLD, command, read, sha, validate_report, write


def guard():
    frozen = read(P / 'freeze.json')
    for name, digest in frozen.items():
        path = ROOT / name
        if path.exists():
            assert sha(path) == digest, name
        else:
            cleanup = read(P / 'cleanup.json')
            assert cleanup['removed'] is True
            assert any(b['path'] == str(path) and b['sha256'] == digest for b in cleanup['binaries']), name
    source = read(P / 'base.json')['source']
    current = {str(f.relative_to(ROOT)) for f in
               list((ROOT/'crates').rglob('*.rs')) + list((ROOT/'crates').rglob('Cargo.toml'))
               + [ROOT/'Cargo.toml', ROOT/'Cargo.lock']}
    assert current == set(source), 'production source census changed'
    for f, h in source.items():
        assert sha(ROOT / f) == h, f
    for f, h in read(P / 'constraints.json').items():
        assert sha(ROOT / f) == h, f


def freeze():
    assert not (P / 'freeze.json').exists()
    assert not (P / 'captures').exists()
    files = [f for f in P.rglob('*') if f.is_file()]
    files += [OLD / 'oracle.json', OLD / 'artifact-manifest.json']
    files += [ROOT / c['path'] for c in read(P / 'cases.json')]
    for variant in ('legacy', 'controls'):
        for binary in read(P / f'{variant}-build.json')['binaries'].values():
            from pathlib import Path
            files.append(Path(binary['path']))
    write(P / 'freeze.json', {str(f.relative_to(ROOT)) if f.is_relative_to(ROOT) else str(f): sha(f)
                              for f in sorted(files)})
    guard()
    print('PASS frozen prospective inputs, source and binary identities')


def capture():
    guard()
    preflight = read(P / 'preflight.json')
    assert preflight['status'] == 'passed' and preflight['freeze_sha256'] == sha(P / 'freeze.json')
    out = P / 'captures'
    assert not out.exists()
    out.mkdir()
    rows = []
    for index, row in enumerate(read(P / 'plan.json')['schedule']):
        cmd = command(row)
        name = f'{index:03d}-{row["lane"]}-{row["case"]}-{row["arm"]}'
        start = time.monotonic_ns()
        with (out / f'{name}.json').open('wb') as stdout, (out / f'{name}.stderr').open('wb') as stderr:
            result = subprocess.run(cmd, cwd=ROOT, stdout=stdout, stderr=stderr)
        end = time.monotonic_ns()
        entry = dict(**row, command=cmd, exit_code=result.returncode, start_monotonic_ns=start,
                     end_monotonic_ns=end, output=f'{name}.json', sha256=sha(out / f'{name}.json'),
                     stderr_sha256=sha(out / f'{name}.stderr'))
        rows.append(entry)
        write(out / 'manifest.json', dict(status='running', freeze_sha256=sha(P / 'freeze.json'),
                                         preflight_sha256=sha(P / 'preflight.json'), runs=rows))
        assert result.returncode == 0
        validate_report(read(out / f'{name}.json'), row)
        print('PASS capture', index, row['lane'], row['case'], row['arm'], flush=True)
    guard()
    write(out / 'manifest.json', dict(status='complete', freeze_sha256=sha(P / 'freeze.json'),
                                     preflight_sha256=sha(P / 'preflight.json'), runs=rows))


if __name__ == '__main__':
    {'freeze': freeze, 'capture': capture}[sys.argv[1]]()
