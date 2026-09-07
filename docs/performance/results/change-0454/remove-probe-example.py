#!/usr/bin/env python3
"""Remove only the retired change-0454 example's exact build artifacts."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
EXAMPLE = 'native_cross_copy_probe_0454'
FINGERPRINT = 'litchi-pptx-6837d8f610535446'
NAMES = [
    f'target/release/examples/{EXAMPLE}',
    f'target/release/examples/{EXAMPLE}.d',
    f'target/release/examples/{EXAMPLE}-6837d8f610535446',
    f'target/release/examples/{EXAMPLE}-6837d8f610535446.d',
    *[f'target/release/.fingerprint/{FINGERPRINT}/{name}' for name in (
        f'example-{EXAMPLE}', f'example-{EXAMPLE}.json',
        f'dep-example-{EXAMPLE}', 'invoked.timestamp',
    )],
]


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    precleanup = ROOT / 'checks/precleanup.json'
    if json.loads(precleanup.read_text()).get('status') != 'pass':
        raise ValueError('completed precleanup verification is required')
    if (REPO / f'crates/litchi-pptx/examples/{EXAMPLE}.rs').exists():
        raise ValueError('the temporary source example still exists')
    rows = []
    for name in NAMES:
        path = REPO / name
        if not path.is_file() or any(parent.is_symlink() for parent in (path, *path.parents)):
            raise ValueError('unexpected example artifact: ' + name)
        raw = path.read_bytes()
        rows.append({'path': name, 'bytes': len(raw), 'sha256': sha(raw)})
    fingerprint = REPO / 'target/release/.fingerprint' / FINGERPRINT
    expected = {REPO / name for name in NAMES if '/.fingerprint/' in name}
    if set(fingerprint.iterdir()) != expected:
        raise ValueError('fingerprint contains another build artifact')
    record = {'status': 'inventoried', 'artifacts': rows,
              'driver_sha256': sha(Path(__file__).read_bytes())}
    with (ROOT / 'checks/retired-example-artifacts.json').open('x') as output:
        output.write(json.dumps(record, indent=2) + '\n')
    for name in NAMES:
        (REPO / name).unlink()
    fingerprint.rmdir()
    if any((REPO / name).exists() for name in NAMES):
        raise ValueError('retired example artifact remains')
    with (ROOT / 'checks/retired-example-removal.json').open('x') as output:
        output.write(json.dumps({'status': 'pass', 'files_removed': len(rows),
            'bytes_removed': sum(row['bytes'] for row in rows),
            'artifact_manifest_sha256': sha((ROOT / 'checks/retired-example-artifacts.json').read_bytes()),
            'driver_sha256': record['driver_sha256']}, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'files_removed': len(rows)}))


if __name__ == '__main__':
    main()
