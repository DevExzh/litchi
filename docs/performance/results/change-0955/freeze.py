"""Recreate the final source census; run with a new output path.

python3 freeze.py SOURCE_SNAPSHOT REPOSITORY BASE_COMMIT NEW_MANIFEST
"""
from pathlib import Path
import hashlib
import json
import subprocess
import sys

source, repository, base, output = sys.argv[1:]
source, repository, output = Path(source), Path(repository), Path(output)
if output.exists():
    raise SystemExit('refusing existing manifest')
owned = {
    'tools/perf-baseline/src/lib.rs',
    'tools/perf-baseline/src/xlsx_managed_compression.rs',
    'tools/perf-baseline/examples/xlsx_managed_compression.rs',
    'tools/perf-baseline/README.md',
}
entries = subprocess.check_output(['git', '-C', str(repository), 'ls-tree', '-rz', base]).split(b'\0')
expected = {}
for entry in entries:
    if not entry:
        continue
    info, name = entry.split(b'\t')
    mode, kind, sha = info.decode().split()
    expected[name.decode()] = sha
records = []
for file in sorted(source.rglob('*')):
    if not file.is_file():
        continue
    name = file.relative_to(source).as_posix()
    data = file.read_bytes()
    if name not in owned:
        assert name in expected, name
        blob = b'blob ' + str(len(data)).encode() + b'\0' + data
        assert hashlib.sha1(blob).hexdigest() == expected[name], name
    else:
        assert data == (repository / name).read_bytes(), name
    records.append({'path': name, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()})
assert set(expected) <= {entry['path'] for entry in records}
with output.open('x') as stream:
    stream.write(json.dumps(records, indent=2) + '\n')
