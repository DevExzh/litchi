#!/usr/bin/env python3
"""Replay a sealed copy and reject independently resealed corruptions."""
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def read(path):
    return json.loads(path.read_text())


def write(path, value):
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def seal(root):
    files = sorted(p for p in root.rglob('*') if p.is_file() and p.name != 'SHA256SUMS')
    assert all(not p.is_symlink() and '__pycache__' not in p.parts for p in files)
    (root / 'SHA256SUMS').write_text(''.join(
        f'{sha(p)}  {p.relative_to(root).as_posix()}\n' for p in files))


def run(root):
    result = subprocess.run([sys.executable, '-B', str(root / 'verify.py')],
                            cwd=root, capture_output=True, text=True)
    return dict(exit_code=result.returncode, stdout=result.stdout, stderr=result.stderr)


def change_report(root, label, mutate):
    report_path = root / 'captures' / f'{label}.report.json'
    report = read(report_path)
    mutate(report)
    write(report_path, report)
    receipt_path = root / 'captures' / f'{label}.json'
    receipt = read(receipt_path)
    receipt['artifacts'][report_path.name] = dict(
        bytes=report_path.stat().st_size, sha256=sha(report_path))
    write(receipt_path, receipt)


def mutate(root, name):
    if name == 'spool_extent':
        change_report(root, 'r1-normal-8-spool',
                      lambda r: r['operations'][0].update(scratch_bytes=577))
    elif name == 'output_digest':
        change_report(root, 'r1-normal-8-control',
                      lambda r: r['operations'][0].update(output_sha256='0' * 64))
    elif name == 'allocator_live_exit':
        def alter(report):
            allocation = report['operations'][0]['allocation']
            allocation['live_bytes_after'] += 1
        change_report(root, 'r1-allocator-8-spool', alter)
    elif name == 'instrumentation':
        change_report(root, 'r1-normal-8-control',
                      lambda r: r.update(instrumentation='system_allocator_operation_scoped'))
    elif name == 'summary_arithmetic':
        path = root / 'summary.json'
        value = read(path)
        value['rows']['r1-normal-8-control']['elapsed_ns']['mean'] += 1
        write(path, value)
    elif name == 'required_gate_omitted':
        path = root / 'rust-validation.json'
        value = read(path)
        value['required'].remove('libraries-tests-opc-guard')
        write(path, value)
    elif name == 'final_gate_exclusion':
        path = root / 'rust-validation.json'
        value = read(path)
        value['attempts']['harness-tests']['source_exclusions'] = [
            'tools/perf-baseline/src/pptx_metadata_spool.rs']
        write(path, value)
    else:
        raise AssertionError(name)
    seal(root)


def main():
    assert (ROOT / 'SHA256SUMS').is_file()
    assert all(not Path(spec['path']).exists()
               for spec in read(ROOT / 'binaries.json').values())
    result = dict(validation_seal_sha256=sha(ROOT / 'SHA256SUMS'),
                  original_runtime_binaries_absent=True, probes={})
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0478-portable-') as temp:
        copy = Path(temp) / 'bundle'
        shutil.copytree(ROOT, copy)
        result['baseline'] = run(copy)
        assert result['baseline']['exit_code'] == 0, result
        for name in ('spool_extent', 'output_digest', 'allocator_live_exit',
                     'instrumentation', 'summary_arithmetic',
                     'required_gate_omitted', 'final_gate_exclusion'):
            shutil.rmtree(copy)
            shutil.copytree(ROOT, copy)
            mutate(copy, name)
            result['probes'][name] = run(copy)
            assert result['probes'][name]['exit_code'] != 0, (name, result)
    result.update(temporary_copy_removed=True, status='pass')
    with (ROOT / 'portable.json').open('x') as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print(json.dumps(result, sort_keys=True))


if __name__ == '__main__':
    main()
