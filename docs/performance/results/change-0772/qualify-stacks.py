"""Account for every sample and reject a profile without the timed caller."""
import gzip
import hashlib
import json
from pathlib import Path
import re

P = Path(__file__).resolve().parent
F = P / 'profile-qualification'
HEADER = re.compile(r'^\S+\s+\d+/\d+\s+[\d.]+:\s+(\d+)\s+cycles:u:\s*$')
ROOT = re.compile(r'litchi_perf_baseline::run_opc_mutated_save\+0x([0-9a-f]+)\b')

if __name__ == '__main__':
    capture = json.loads((F / 'manifest.json').read_text())
    build = json.loads((P / 'build.json').read_text())
    assert capture['exit'] == 0 and capture['binary_sha256'] == build['binary_sha256']
    assert capture['boundary_sha256'] == hashlib.sha256((P / 'timed-boundary.json').read_bytes()).hexdigest()
    storage = json.loads((P / 'profile-storage.json').read_text())
    for name, digest in capture['files'].items():
        if name == storage['raw_path']:
            packed = (P / storage['archive']).read_bytes()
            assert hashlib.sha256(packed).hexdigest() == storage['encoded_sha256']
            raw = gzip.decompress(packed)
            assert len(raw) == storage['decoded_bytes']
            assert hashlib.sha256(raw).hexdigest() == digest == storage['decoded_sha256']
        else:
            assert hashlib.sha256((P / name).read_bytes()).hexdigest() == digest, name
    report = json.loads((F / 'report.json').read_text())
    assert report['binary_identity']['binary_sha256'] == build['binary_sha256']
    assert report['environment']['cpu_affinity'] == '12'
    assert report['environment']['git_revision'] == json.loads((P / 'source.json').read_text())['head']
    assert report['configuration']['samples_per_case'] == 30
    assert report['configuration']['warmup_iterations_per_case'] == 0
    assert len(report['results']) == 1
    measured = report['results'][0]
    assert measured['case'] == 'opc_mutated_save'
    expected = json.loads((P / 'descriptive-projection.json').read_text())['opc_mutated_save/few-large-incompressible']
    assert {key: measured.get(key) for key in ('corpus', 'sink', 'output_sha256')} == expected
    assert len(measured['elapsed_ns']['samples']) == 30
    receipt = json.loads((F / 'decode.json').read_text())
    archived = (P / receipt['output']).read_bytes()
    assert hashlib.sha256(archived).hexdigest() == receipt['encoded_sha256']
    data = gzip.decompress(archived)
    assert hashlib.sha256(data).hexdigest() == receipt['decoded_sha256']
    assert len(data) == receipt['decoded_bytes']
    buckets = {key: {'samples': 0, 'period': 0} for key in
               ('exact_timed_caller', 'other_root_offset', 'no_recognized_root')}
    offsets = {}
    for block in data.decode().strip().split('\n\n'):
        lines = block.splitlines()
        match = HEADER.fullmatch(lines[0])
        assert match, lines[0]
        period = int(match.group(1))
        roots = ROOT.findall(block)
        assert len(roots) <= 1, roots
        if roots:
            offset = int(roots[0], 16)
            offsets[hex(offset)] = offsets.get(hex(offset), 0) + 1
            # perf may resolve a caller's return IP after subtracting one.
            key = 'exact_timed_caller' if offset in (0x7d1, 0x7d2) else 'other_root_offset'
        else:
            key = 'no_recognized_root'
        buckets[key]['samples'] += 1
        buckets[key]['period'] += period
    strict = buckets['exact_timed_caller']['samples']
    result = {'qualification': 'rejected' if strict == 0 else 'requires_further_review',
              'buckets': buckets, 'root_offsets': offsets,
              'samples': sum(row['samples'] for row in buckets.values()),
              'period': sum(row['period'] for row in buckets.values()),
              'decoded_sha256': receipt['decoded_sha256'],
              'limitation': 'No recognized root does not mean outside the timed region. '
                            'Missing/truncated unwind context remains unclassified. '
                            'No Deflate fraction or removable-cost inference is supported.'}
    (F / 'qualification.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(result))
