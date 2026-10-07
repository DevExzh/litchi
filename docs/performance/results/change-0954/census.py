"""Verify retained exporter artifacts and model advisory Deflate wave widths.

Run from repository root. No benchmark workload or production code is executed.
"""
import argparse
import hashlib
import json
from pathlib import Path
import struct
import zipfile


def digest(data):
    return hashlib.sha256(data).hexdigest()


def describe(path):
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data), 'sha256': digest(data)}


def checked(root, record):
    data = (root / record['path']).read_bytes()
    assert len(data) == record['bytes'] and digest(data) == record['sha256'], record
    return data


def compressed(archive, info):
    archive.fp.seek(info.header_offset)
    header = archive.fp.read(30)
    assert header[:4] == b'PK\x03\x04'
    name, extra = struct.unpack_from('<HH', header, 26)
    archive.fp.seek(name + extra, 1)
    return archive.fp.read(info.compress_size)


def census(root):
    manifest_path = root / 'manifest.json'
    manifest = json.loads(manifest_path.read_text())
    prior_path = Path('docs/performance/results/change-0953/task-census.json')
    prior = {r['shape']: r for r in json.loads(prior_path.read_text())['reports']}
    cases = []
    for case in manifest['cases']:
        source = checked(root, case['source'])
        output = checked(root, case['output'])
        if case['case'] == 'one-percent':
            assert digest(source) == prior[case['shape']]['archive_sha256']
            assert digest(output) == prior[case['shape']]['output_sha256']
        rows = []
        with zipfile.ZipFile(root / case['source']['path']) as src, zipfile.ZipFile(root / case['output']['path']) as dst:
            assert src.namelist() == dst.namelist()
            assert len(set(dst.namelist())) == len(dst.namelist())
            changed = {name for name in src.namelist() if src.read(name) != dst.read(name)}
            assert changed == {p['member'] for p in case['parts']}
            untouched = set(src.namelist()) - changed
            for name in untouched:
                assert compressed(src, src.getinfo(name)) == compressed(dst, dst.getinfo(name))
            for part in case['parts']:
                name = part['member']
                original = checked(root, part['original'])
                replacement = checked(root, part['replacement'])
                assert src.read(name) == original and dst.read(name) == replacement
                info = dst.getinfo(name)
                assert info.compress_type == zipfile.ZIP_DEFLATED
                assert info.file_size == len(replacement)
                rows.append({'member': name, 'logical_bytes': len(replacement),
                             'compressed_bytes': info.compress_size,
                             'replacement_sha256': digest(replacement)})
        total = sum(r['logical_bytes'] for r in rows)
        largest = max(r['logical_bytes'] for r in rows)
        remainder = total - largest
        widths = []
        for workers in (1, 2, 4, 8):
            widths.append({'workers': workers, 'max_in_flight_tasks': workers,
                           'cold_width': min(workers, len(rows)) if remainder >= 65536 else 1,
                           'warm_width': min(workers, len(rows)) if remainder >= 8192 else 1})
        cases.append({'id': case['id'], 'source': case['source'], 'output': case['output'],
                      'matches_0953_identity': case['case'] == 'one-percent',
                      'tasks': rows, 'task_count': len(rows), 'total_bytes': total,
                      'largest_bytes': largest, 'remainder_bytes': remainder,
                      'untouched_compressed_payloads_verified': len(untouched),
                      'advisory_widths': widths})
    return {'schema': 'litchi.performance.0954.task-census.v1',
            'manifest': describe(manifest_path), 'prior_identity': describe(prior_path),
            'policy': {'cold_remainder_floor': 65536, 'warm_remainder_floor': 8192,
                       'min_task_bytes': 0,
                       'limits': 'Advisory only; actual execution also requires explicit context, memory and task budgets. Warm means an existing pool/session, not warm CPU caches.'},
            'cases': cases}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('artifacts', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    result = census(args.artifacts)
    with args.output.open('x') as out:
        json.dump(result, out, indent=2, sort_keys=True)
        out.write('\n')
