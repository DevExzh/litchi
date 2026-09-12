"""Replay sealed historical planning evidence and attribute transient attributes."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PRIOR = HERE.with_name('change-0530')

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def seal(folder, expected):
    assert sha(folder / 'SHA256SUMS') == expected
    entries = {}
    for line in (folder / 'SHA256SUMS').read_text().splitlines():
        digest, name = line.split('  ', 1)
        assert name not in entries and not Path(name).is_absolute()
        assert '..' not in Path(name).parts and name != 'SHA256SUMS'
        entries[name] = digest
    actual = {str(f.relative_to(folder)): sha(f) for f in folder.rglob('*')
              if f.is_file() and f.name != 'SHA256SUMS'}
    assert entries == actual
    return dict(entries=len(entries), sha256=expected)

def analyze():
    plan = json.loads((HERE / 'plan.json').read_text())
    seals = {name: seal(HERE.with_name(name), digest)
             for name, digest in plan['prior_seals'].items()}
    assert sha(HERE / 'source-binding.json') == plan['source_binding_sha256']
    source = json.loads((HERE / 'source-binding.json').read_text())
    old = json.loads((PRIOR / 'baseline/source-manifest.json').read_text())
    for name, digest in source.items():
        assert sha(REPO / name) == digest == old[name], name
    spec = importlib.util.spec_from_file_location('prior_planning', PRIOR / 'analyze_planning.py')
    previous = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(previous)
    replay = previous.analyze(False)
    assert replay == json.loads((PRIOR / 'planning-analysis.json').read_text())
    helper = previous.h
    rows = []
    for record in replay['rows']:
        stem = PRIOR / 'baseline' / f"profile-r{record['repeat']}-{record['shape']}"
        inclusive = Path(str(stem) + '.inclusive.txt').read_text()
        exclusive = Path(str(stem) + '.self.txt').read_text()
        raw = PRIOR / record['selected']
        total = record['parts'][-1]['summary_ir']
        owners = {}
        for name in plan['selected_owners']:
            inc = helper.parse_annotation(inclusive, name, str(stem))
            own = helper.parse_annotation(exclusive, name, str(stem))
            children = helper.direct_map(inc['direct'])
            assert children == helper.direct_map(own['direct'])
            assert inc['selected_ir'] == own['selected_ir'] + sum(children.values())
            for child, cost in children.items():
                edge = helper.target_edge_summary(raw, child, name)
                assert edge['inclusive_ir'] == cost
            owners[name] = dict(inclusive_ir=inc['selected_ir'], self_ir=own['selected_ir'],
                                direct=children, planning_share_percent=100*inc['selected_ir']/total)
        scan, decode = plan['selected_owners']
        edge = helper.target_edge_summary(raw, decode, scan)
        assert edge['inclusive_ir'] == owners[scan]['direct'][decode]
        rows.append(dict(shape=record['shape'], repeat=record['repeat'], planning_ir=total,
                         raw_sha256=sha(raw), owners=owners))
    return dict(status='pass', plan_sha256=sha(HERE/'plan.json'), seals=seals,
                source_binding_sha256=sha(HERE/'source-binding.json'), rows=rows,
                scope='Historical0530 selected planning dumps; current relevant source matches. '
                      'Nested owner totals overlap. Callee Ir is not operation-local allocation counts. '
                      'No fresh timings or accepted optimization.')

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=HERE/'analysis.json')
    args = parser.parse_args()
    result = analyze()
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True)+'\n')
    print('Historical descendant/source replay PASS; four selected planning dumps.')
