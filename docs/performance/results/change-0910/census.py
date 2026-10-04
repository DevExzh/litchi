"""Run independent raw ZIP checks on the first representative output of each route."""
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
HELPER = ROOT / 'docs/performance/results/change-0909/zip-census.py'


def descriptor(path):
    data = path.read_bytes()
    return {'path': str(path.relative_to(ROOT)), 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}


def main():
    assert not (P / 'census-runs.json').exists()
    assert descriptor(HELPER)['sha256'] == '0cdee4682959e9a448d81bc598a810c7c16f780bf378bcd8faae294f1cf1efd3'
    runs = []
    selected = sorted(f'ppt/slides/slide{n + 1}.xml' for n in [0, 28, 57, 85, 114, 142, 171, 199])
    for route in ('materialized', 'owned'):
        folder = P / 'capture' / f'block-00-{route}' / 'artifacts'
        report = P / f'census-{route}.json'
        argv = ['python3', '-B', str(HELPER), '--source', str(folder / 'source.pptx'), '--output', str(folder / 'output.pptx'), '--report', str(report), '--expected-generated-slides']
        result = subprocess.run(argv, cwd=ROOT, capture_output=True, text=True)
        assert result.returncode == 0, result.stderr
        census = json.loads(report.read_text())
        assert census['status'] == 'pass' and not census['violations']
        assert len(census['members']) == 445 and len(census['unchanged_members']) == 437
        assert sorted(census['changed_members']['names']) == selected
        assert census['ordered_filenames']['equal'] and census['archive']['comment_comparison']['equal']
        physical = {side: [r['name'] for r in sorted(census['members'], key=lambda r: r[side]['physical']['local_header_offset'])] for side in ('source', 'output')}
        assert physical['source'] == physical['output']
        runs.append({'label': route, 'argv': argv, 'exit_code': result.returncode, 'stdout': result.stdout, 'stderr': result.stderr, 'helper': descriptor(HELPER), 'report': descriptor(report), 'physical_order_equal': True, 'changed_members': selected, 'unchanged_members': 437})
    with (P / 'census-runs.json').open('x') as stream:
        json.dump({'status': 'pass', 'helper': descriptor(Path(__file__)), 'runs': runs}, stream, indent=2)
        stream.write('\n')
    print('0910 independent raw census PASS for both routes')


if __name__ == '__main__':
    main()
