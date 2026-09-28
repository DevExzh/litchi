"""Independent sample-block census for the retained, exact-owner perf frames."""
from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys

P = Path(__file__).resolve().parent
OWNER = 'pptx_edit_profile_0822::edit_region_0822'
HEADER = re.compile(r'^\S+\s+\d+\s+\d+\.\d+:\s+(\d+)\s+cycles:u:\s*$')
FRAME = re.compile(r'^\s*([0-9a-fA-F]+)\s+(.+)\s+\((.+)\)\s*$')


def main():
    assert sys.argv[1:] in (['--write'], ['--check'])
    complete = json.loads((P / 'perf/decode-complete.json').read_text())
    result = {'schema': 'litchi.performance.0822.frame-audit.v1',
              'status': complete['status'], 'repeats': []}
    if complete['status'] == 'available':
        binary = json.loads((P / 'build.json').read_text())['binaries']['fp']['artifact']
        members = json.loads((P / 'perf/compression.json').read_text())
        for member in members:
            if member['kind'] != 'frames':
                continue
            stored = Path(member['compressed']['path']).read_bytes()
            assert len(stored) == member['compressed']['bytes']
            assert hashlib.sha256(stored).hexdigest() == member['compressed']['sha256']
            raw = gzip.decompress(stored)
            assert len(raw) == member['original']['bytes']
            assert hashlib.sha256(raw).hexdigest() == member['original']['sha256']
            stacks = []
            current = None
            for line in raw.decode().splitlines():
                header = HEADER.fullmatch(line)
                if header:
                    current = {'period': int(header[1]), 'frames': []}
                    stacks.append(current)
                elif line.strip():
                    frame = FRAME.fullmatch(line)
                    assert frame and current is not None, repr(line)
                    function = re.sub(r'\+0x[0-9a-fA-F]+$', '', frame[2])
                    current['frames'].append((function, frame[3]))
                else:
                    current = None
            leaves = Counter()
            selected = self_leaves = unknown_interiors = unknown_leaves = 0
            owner_period = 0
            for stack in stacks:
                frames = stack['frames']
                assert frames
                owners = [i for i, f in enumerate(frames) if f == (OWNER, binary['path'])]
                assert len(owners) <= 1
                if not owners:
                    continue
                selected += 1
                owner_period += stack['period']
                index = owners[0]
                self_leaves += int(index == 0)
                leaves[frames[0]] += 1
                unknown_leaves += int('[unknown]' in frames[0])
                unknown_interiors += int(any('[unknown]' in f for f in frames[1:index]))
            assert selected > 0 and sum(leaves.values()) == selected
            result['repeats'].append({
                'repeat': member['repeat'], 'whole_samples': len(stacks),
                'owner_samples': selected, 'outside_owner_samples': len(stacks) - selected,
                'owner_period': owner_period, 'owner_self_leaves': self_leaves,
                'unknown_leaf_samples': unknown_leaves,
                'unknown_interior_samples': unknown_interiors,
                'leaves': [{'symbol': f, 'dso': dso, 'samples': n}
                           for (f, dso), n in sorted(leaves.items(), key=lambda x: (-x[1], x[0]))],
            })
        assert len(result['repeats']) == 2
    encoded = json.dumps(result, sort_keys=True, indent=2) + '\n'
    out = P / 'frame-audit.json'
    if sys.argv[1] == '--write':
        assert not out.exists()
        out.write_text(encoded)
    else:
        assert out.read_text() == encoded
    print('0822 independent frame audit PASS')


if __name__ == '__main__':
    main()
