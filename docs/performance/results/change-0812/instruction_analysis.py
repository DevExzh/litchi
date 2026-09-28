"""Join independently parsed fresh leaf offsets to their exact capture binary."""
import collections
import gzip
import json
from pathlib import Path
import re
import sys
import inputs as c

OWNER = 'namespace_uri_probe::capture_region_0793'
SYMBOL = 'litchi_pptx::notes::codec::scan_processed_xml'

def checked(row):
    assert c.artifact(row['path']) == row, row['path']
    return Path(row['path'])

def analyze():
    custody = c.verify()
    rebuilt = c.read(c.P / 'rebuild/receipt.json')
    expected = rebuilt['binary']
    assert rebuilt['status'] == 'binary-mismatch' and not rebuilt['historical_mapping_authorized']
    assert not rebuilt['binary_comparison']['exact'] and rebuilt['assembly'] is None
    checked(rebuilt['prefreeze']); checked(rebuilt['commands'])
    prefreeze = c.read(rebuilt['prefreeze']['path'])
    for key in ('driver', 'plan', 'manifest', 'source_before_artifact'):
        checked(prefreeze[key])
    assert rebuilt['build']['exit_code'] == 0
    checked(rebuilt['build']['log'])
    assert c.read(checked(rebuilt['source']['before'])) == c.read(checked(rebuilt['source']['after']))
    fresh = c.P / 'fresh'
    completed = c.read(fresh / 'complete.json')
    assert completed['binary'] == expected
    assert completed['reports'] == 2 and completed['samples'] == 200
    base, size = completed['address'], completed['size']
    asm = {'symbol': SYMBOL, 'mangled_symbol': completed['symbol'], 'address': base, 'size': size,
           'objdump': completed['assembly'], 'symbols': c.artifact(fresh / 'symbols-resumed.txt')}
    nm = [line.split() for line in checked(asm['symbols']).read_text().splitlines()
          if line.endswith(' ' + asm['mangled_symbol'])]
    assert len(nm) == 1 and int(nm[0][0],16) == base and int(nm[0][1],16) == size
    assert nm[0][2].lower() == 't'
    instructions = {}
    for line in checked(asm['objdump']).read_text().splitlines():
        m = re.fullmatch(r'\s*([0-9a-f]+):\s+(.+)', line)
        if m:
            offset = int(m[1],16) - base
            assert 0 <= offset < size and offset not in instructions
            instructions[offset] = m[2]
    assert instructions
    rows = []
    frames = [v for v in c.read(fresh / 'compression.json') if v['kind'] == 'frames']
    assert [v['repeat'] for v in frames] == [0,1]
    for descriptor in frames:
        raw = gzip.decompress(checked(descriptor['compressed']).read_bytes())
        assert len(raw) == descriptor['original']['bytes']
        assert __import__('hashlib').sha256(raw).hexdigest() == descriptor['original']['sha256']
        counts, periods = collections.Counter(), collections.Counter()
        whole = owner = unknown = 0
        for block in raw.decode().strip().split('\n\n'):
            lines = block.splitlines()
            header = re.fullmatch(r'\S+\s+\d+\s+(\d+\.\d+):\s+(\d+) cycles:u:\s*', lines[0])
            assert header, lines[0]
            whole += 1
            stack = []
            for line in lines[1:]:
                match = re.fullmatch(r'\s*[0-9a-f]+ (.+?)(?:\+0x([0-9a-f]+))? \((.+)\)',line)
                assert match, line
                stack.append((match[1], int(match[2] or '0',16), match[3]))
            names = [frame[0] for frame in stack]
            if OWNER not in names:
                continue
            assert names.count(OWNER) == 1
            owner += 1
            unknown += any('[unknown]' in n or n == '??' for n in names[:names.index(OWNER)])
            if stack[0][0] != SYMBOL:
                continue
            name, offset, dso = stack[0]
            assert dso == expected['path'] and offset in instructions
            counts[offset] += 1
            periods[offset] += int(header[2])
        rows.append({'repeat': descriptor['repeat'], 'whole_samples': whole, 'owner_samples': owner,
                     'unknown_interior': unknown, 'leaf_samples': sum(counts.values()), 'leaf_period': sum(periods.values()),
                     'offsets': [{'offset': offset, 'offset_hex': hex(offset), 'samples': count,
                                  'period': periods[offset], 'instruction': instructions[offset]}
                                 for offset,count in sorted(counts.items())]})
    oracle = c.read(c.OLD / 'perf/0.json')
    reports = []
    for receipt in c.read(fresh / 'receipts.json'):
        report = c.read(checked(receipt['report']))
        assert receipt['binary'] == expected and receipt['exit_code'] == 0
        for key in ('schema', 'tool', 'source', 'fixture', 'slides', 'shapes_per_slide',
                    'marker', 'timing_scope', 'mode', 'shape', 'warmup'):
            assert report[key] == oracle[key], key
        assert len(report['samples']) == 100 and report['warmup'] == 0
        for index, sample in enumerate(report['samples']):
            assert sample['index'] == index
            for key in ('source_sha256', 'output', 'verification'):
                assert sample[key] == oracle['samples'][0][key], key
            assert sample['elapsed_ns'] == sample['metrics']['elapsed_ns'] > 0
        text = checked(receipt['output']).read_text()
        assert not __import__('re').search(r'\blost\b', text, __import__('re').I)
        rows[receipt['repeat']]['lost_event_lines'] = 0
        reports.append(receipt['report'])
    return {'schema': 'litchi.performance.0812.instruction-analysis.v1', 'inputs': custody,
            'binary': expected, 'assembly': asm, 'rows': rows, 'reports': reports,
            'measured_outputs_verified': 200, 'historical_offsets_mapped': False,
            'scope': 'Fresh exact-binary sample localization; sampling skid and frame-pointer perturbation apply. No causal instruction cost or ordinary-build phase fractions.'}

if __name__ == '__main__':
    result = analyze()
    encoded = json.dumps(result,indent=2,sort_keys=True)+'\n'
    output = c.P / 'instruction-analysis.json'
    if '--check' in sys.argv:
        assert output.read_text() == encoded
    else:
        assert sys.argv[1:] == ['--write'] and not output.exists()
        output.write_text(encoded)
    print('0812 exact-binary instruction localization PASS')
