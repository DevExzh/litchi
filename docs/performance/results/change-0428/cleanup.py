#!/usr/bin/env python3
"""Remove only the hash-bound executable copied for this task's capture."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    build = json.loads((ROOT / 'build.json').read_text())
    temporary = Path('/tmp/litchi-goal-0428-binaries')
    binary = Path(build['capture_binary'])
    assert binary.parent == temporary
    assert list(temporary.iterdir()) == [binary]
    assert digest(binary) == build['binary_sha256']
    original = REPO / build['original_binary']
    assert digest(original) == build['binary_sha256']
    binary.unlink()
    temporary.rmdir()
    record = {'status': 'pass', 'removed_directory': str(temporary),
              'removed_binary_sha256': build['binary_sha256'],
              'temporary_directory_absent': not temporary.exists(),
              'original_build_binary_retained': original.is_file(),
              'root_target_retained': (REPO / 'target').is_dir(),
              'harness_target_retained': (REPO / 'tools/perf-baseline/target').is_dir()}
    assert all(record[key] for key in ['temporary_directory_absent', 'original_build_binary_retained',
                                       'root_target_retained', 'harness_target_retained'])
    (ROOT / 'cleanup.json').write_text(json.dumps(record, indent=2) + '\n')
    print(json.dumps(record))


if __name__ == '__main__':
    main()
