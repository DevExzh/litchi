"""Join exact-owner sampled leaf offsets to the actual fp assembly, offline."""
import collections
import gzip
import hashlib
import json
import re
import sys
from pathlib import Path
import custody as c

OWNER = 'namespace_uri_probe::capture_region_0793'
SYMBOLS = {'scanner': 'litchi_pptx::notes::codec::scan_processed_xml',
           'inspector': 'litchi_pptx::notes::codec::inspect_element'}

def analyze():
    build = c.read(c.P / 'build/build.json')
    binary = build['binaries']['fp']
    receipt = c.read(c.P / 'assembly/receipt.json')
    assemblies = {r['name']: r for r in receipt['rows'] if r['variant'] == 'fp'}
    assert set(assemblies) == set(SYMBOLS)
    instructions = {}
    for name, row in assemblies.items():
        assert row['binary'] == binary and row['source'] == build['source']
        artifact = row['objdump']['assembly']
        assert c.artifact(artifact['path']) == artifact
        values = {}
        for line in Path(artifact['path']).read_text().splitlines():
            match = re.fullmatch(r'\s*([0-9a-f]+):\s+(.+)', line)
            if match:
                offset = int(match[1], 16) - row['address']
                assert 0 <= offset < row['size']
                values[offset] = match[2]
        assert values
        instructions[name] = values
    frames = c.read(c.P / 'perf/frame-receipts.json')
    decodes = c.read(c.P / 'perf/decode-receipts.json')
    assert [r['repeat'] for r in frames] == [0, 1]
    assert [r['repeat'] for r in decodes] == [0, 1]
    reports = []
    for frame, decode in zip(frames, decodes):
        assert frame['repeat'] == decode['repeat']
        assert decode['binary'] == binary and frame['original'] == decode['frames']
        assert c.artifact(frame['compressed']['path']) == frame['compressed']
        raw = gzip.decompress(Path(frame['compressed']['path']).read_bytes())
        assert len(raw) == frame['original']['bytes']
        assert hashlib.sha256(raw).hexdigest() == frame['original']['sha256'] == frame['decompressed_sha256']
        counts = {name: collections.Counter() for name in SYMBOLS}
        periods = {name: collections.Counter() for name in SYMBOLS}
        owner_count = 0
        for block in raw.decode().strip().split('\n\n'):
            lines = block.splitlines()
            header = re.fullmatch(r'\S+\s+\d+\s+(\d+\.\d+):\s+(\d+) cycles:u:\s*', lines[0])
            assert header
            stack = []
            for line in lines[1:]:
                match = re.fullmatch(r'\s*([0-9a-f]+) (.+?)(?:\+0x([0-9a-f]+))? \((.+)\)', line)
                assert match, line
                stack.append((match[2], int(match[3] or '0', 16), match[4]))
            owners = [i for i, item in enumerate(stack) if item[0] == OWNER and item[2] == binary['path']]
            if not owners:
                continue
            assert len(owners) == 1
            owner_count += 1
            if owners[0] == 0:
                continue
            symbol, offset, dso = stack[0]
            for name, expected in SYMBOLS.items():
                if symbol == expected:
                    assert dso == binary['path']
                    assert 0 <= offset < assemblies[name]['size']
                    counts[name][offset] += 1
                    periods[name][offset] += int(header[2])
        symbols = {}
        for name in SYMBOLS:
            rows = [{'offset': offset, 'hex_offset': hex(offset), 'samples': count,
                     'period': periods[name][offset], 'instruction': instructions[name].get(offset),
                     'instruction_boundary': offset in instructions[name]}
                    for offset, count in counts[name].most_common()]
            symbols[name] = {'symbol': SYMBOLS[name], 'symbol_bytes': assemblies[name]['size'],
                             'leaf_samples': sum(counts[name].values()), 'offsets': rows}
        mapped = sum(row['leaf_samples'] for row in symbols.values())
        assert mapped <= owner_count
        reports.append({'repeat': frame['repeat'], 'owner_samples': owner_count,
                        'other_owner_leaf_samples': owner_count - mapped,
                        'frames': frame['compressed'], 'symbols': symbols})
    return {'schema': 'litchi.performance.0814.offset-audit.v1', 'binary': binary,
            'assembly_receipt': c.artifact(c.P / 'assembly/receipt.json'), 'reports': reports,
            'scope': 'Exact fp binary scanner/inspector leaf offsets only; other owner leaves are counted separately. Sampling skid and frame-pointer perturbation prevent causal cycle or ordinary-build phase claims.'}

result = analyze()
out = c.P / 'offset-audit.json'
encoded = json.dumps(result, indent=2, sort_keys=True) + '\n'
if '--check' in sys.argv:
    assert out.read_text() == encoded
else:
    assert sys.argv[1:] == ['--write'] and not out.exists()
    out.write_text(encoded)
print('0814 exact-binary leaf-offset replay PASS')
