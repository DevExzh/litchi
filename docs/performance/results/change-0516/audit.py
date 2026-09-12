"""Read-only receipt and source-custody checks for recorded 0516 epochs.

This module checks the evidence that exists. The final verifier separately
requires the complete lane set for the recorded admission decision.
"""
import datetime
import json
from pathlib import Path

from run import HERE, REPO, SCRATCH, sha


def read(path):
    return json.loads(path.read_text())


def audit():
    epochs = {}
    for path in sorted(HERE.glob('*/source-manifest.json')):
        manifest = read(path)
        assert manifest and all(len(value) == 64 for value in manifest.values())
        epochs[path.parent.name] = sha(path)
    assert epochs.get('before') == sha(HERE.parent / 'change-0515/source-manifest.json')
    failures = read(HERE / 'expected-failures.json') if (HERE / 'expected-failures.json').exists() else {}
    receipts = []
    intervals = []
    for path in sorted(HERE.glob('*/*-receipt.json')):
        receipt = read(path)
        if 'source_manifest_sha256' not in receipt:
            continue
        epoch = path.parent.name
        assert receipt['source_manifest_sha256'] == epochs[epoch], path
        expected_exit = failures.get(str(path.relative_to(HERE)), {}).get('exit_code', 0)
        assert receipt['exit_code'] == expected_exit and receipt['source_unchanged'], path
        artifacts = receipt.get('artifacts', {})
        for name, expected in artifacts.items():
            target = path.parent / name
            assert target.parent == path.parent and sha(target) == expected, target
        if 'log_sha256' in receipt:
            target = path.with_name(path.name.replace('-receipt.json', '.log'))
            assert sha(target) == receipt['log_sha256'], target
        probe = receipt.get('probe')
        if probe:
            assert probe in ('guard', 'fallback'), path
            manifest_path = HERE / (probe + '-source-manifest.json')
            assert receipt['guard_manifest_sha256'] == sha(manifest_path), path
            for name, expected in read(manifest_path).items():
                assert sha(REPO / name) == expected, name
        command = receipt['command']
        if 'active_source_roles' in receipt:
            assert receipt['active_source_roles'], path
            assert set(receipt['active_source_roles']) <= epochs.keys(), path
        start = datetime.datetime.fromisoformat(receipt['started_utc'])
        duration = receipt['elapsed_seconds']
        assert duration > 0 and start.tzinfo is not None, path
        intervals.append((start, start + datetime.timedelta(seconds=duration), str(path.relative_to(HERE))))
        if 'binary_sha256' in receipt:
            if probe:
                allocator = 'allocator' in path.name
                binary = SCRATCH / epoch / (probe + ('-alloc' if allocator else ''))
            else:
                allocator = 'allocator' in path.name
                binary = SCRATCH / epoch / ('litchi-perf-baseline' + ('-alloc' if allocator else ''))
            if binary.exists():
                assert sha(binary) == receipt['binary_sha256'], path
            if 'taskset' in command:
                position = command.index('taskset')
                assert command[position:position + 3] == ['taskset', '-c', '2'], path
                build_name = (probe + '-' if probe else '') + ('build-allocator' if probe and allocator else 'allocator-build' if allocator else 'build')
                built = read(path.parent / (build_name + '-receipt.json'))
                assert receipt['binary_sha256'] == built['binary_sha256'], path
        receipts.append(str(path.relative_to(HERE)))
    intervals.sort()
    for earlier, later in zip(intervals, intervals[1:]):
        assert earlier[1] <= later[0], (earlier[2], later[2], 'overlapping recorded operations')
    return {'epochs': epochs, 'receipt_count': len(receipts), 'receipts': receipts,
            'recorded_operations_serial': True, 'recorded_expected_failures': failures,
            'scope': 'Existing receipts only; lane completeness and semantic metrics are verified separately.'}


if __name__ == '__main__':
    print(json.dumps(audit(), indent=2))
