#!/usr/bin/env python3
"""Separate allocator-instrumented children; their times are not native evidence."""
import json
import subprocess
import time

from build import P, ROOT, TARGET, census, sha

def main():
    expected = json.loads((P / 'source-manifest.json').read_text())
    assert census() == expected
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
               'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics',
               '--bin', 'litchi-perf-baseline-alloc', '--target-dir', str(TARGET), '-j', '2']
    start = time.monotonic()
    with (P / 'build-alloc.log').open('w') as log:
        r = subprocess.run(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
    assert census() == expected
    binary = TARGET / 'release/litchi-perf-baseline-alloc'
    receipt = dict(command=command, exit_code=r.returncode, seconds=time.monotonic()-start,
                   binary=str(binary), binary_sha256=sha(binary) if r.returncode == 0 else None,
                   source_manifest_sha256=sha(P / 'source-manifest.json'))
    (P / 'build-alloc.json').write_text(json.dumps(receipt, indent=2) + '\n')
    assert r.returncode == 0
    for shape in ['medium', 'dense-sparse']:
        for repeat in [1, 2]:
            name = f'alloc-r{repeat}-{shape}'
            command = ['taskset', '-c', '12', str(binary), '--warmup', '2', '--samples', '3',
                       '--case', 'xlsx_source_backed_cell_values_one_percent_edit_save',
                       '--xlsx-cell-crud-shape', shape, '--json', str(P / (name + '.json'))]
            assert not (P / (name + '.receipt.json')).exists()
            with (P / (name + '.stdout')).open('w') as out, (P / (name + '.stderr')).open('w') as err:
                r = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
            assert census() == expected
            assert sha(binary) == receipt['binary_sha256']
            (P / (name + '.receipt.json')).write_text(json.dumps(dict(
                command=command, exit_code=r.returncode, binary_sha256=sha(binary),
                artifacts={name+s:sha(P/(name+s)) for s in ['.json', '.stdout', '.stderr']},
                script_sha256=sha(P / 'allocations.py'),
                source_manifest_sha256=sha(P / 'source-manifest.json')), indent=2)+'\n')
            assert r.returncode == 0
            print(name, 'passed', flush=True)

if __name__ == '__main__':
    main()
