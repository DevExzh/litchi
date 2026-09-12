"""Recompute scanner instruction attribution from the sealed 0525 profiles."""
import argparse
import importlib.util
import json
from pathlib import Path
from verify import HERE, PRIOR, REPO, check_seal, require, sha

HELPER = HERE.parent / 'change-0521/analyze_profiles.py'
spec = importlib.util.spec_from_file_location('retained_0521_profile_helpers', HELPER)
helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(helper)
OWNER = 'litchi_xlsx::cell_values::source::MultiSourceEdit::commit'
REWRITE = 'litchi_xlsx::raw::worksheet::edit::package::rewrite_value_only_with_provenance'
SCAN = 'litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::scan_with_limit'
START = 'litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::start_cell'
ADDRESS = 'litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::cell_address'
SELECTED = (OWNER, REWRITE, SCAN, START, ADDRESS)

def analyze():
    require(check_seal(PRIOR) == 550, 'retained evidence inventory differs')
    helpers = [HELPER, HERE.parent / 'change-0519/analyze_profiles.py',
               HERE.parent / 'change-0519/compare_profile_lanes.py']
    binding = json.loads((HERE / 'source-binding.json').read_text())
    for name, digest in binding['source_files'].items():
        require(sha(REPO / name) == digest, f'current source differs: {name}')
    rows = []
    totals = {name: {'inclusive_ir': 0, 'self_ir': 0, 'direct': {}} for name in SELECTED}
    for repeat in (1, 2):
        for shape in ('medium', 'dense-sparse'):
            stem = PRIOR / 'candidate' / f'profile-r{repeat}-{shape}'
            raw = Path(str(stem) + '.callgrind.4')
            inc_path = Path(str(stem) + '.inclusive.txt')
            self_path = Path(str(stem) + '.self.txt')
            inc_text, self_text = inc_path.read_text(), self_path.read_text()
            owners = {}
            for name in SELECTED:
                inc = helper.parse_annotation(inc_text, name, str(inc_path))
                own = helper.parse_annotation(self_text, name, str(self_path))
                require(inc['direct'] == own['direct'], f'annotation children differ: {name}')
                direct = helper.direct_map(inc['direct'])
                require(inc['selected_ir'] == own['selected_ir'] + sum(direct.values()),
                        f'self plus direct costs differ: {name}')
                for child, cost in direct.items():
                    edge = helper.target_edge_summary(raw, child, name)
                    require(edge['inclusive_ir'] == cost, f'raw direct edge differs: {name} -> {child}')
                owner = {'inclusive_ir': inc['selected_ir'], 'self_ir': own['selected_ir'],
                         'direct': dict(sorted(direct.items(), key=lambda item: (-item[1], item[0])))}
                owners[name] = owner
                totals[name]['inclusive_ir'] += owner['inclusive_ir']
                totals[name]['self_ir'] += owner['self_ir']
                for child, cost in direct.items():
                    totals[name]['direct'][child] = totals[name]['direct'].get(child, 0) + cost
            incoming = helper.target_edge_summary(raw, OWNER, 'litchi_perf_baseline::run_xlsx_cell_values_edit_save')
            require(incoming['positive_edge_count'] == 1 and incoming['calls'] == 1,
                    'selected commit does not have one direct measured runner call')
            require(incoming['inclusive_ir'] == owners[OWNER]['inclusive_ir'] ==
                    helper.summary_ir(raw.read_text(), str(raw)), 'selected owner differs from raw summary')
            require(owners[OWNER]['direct'][REWRITE] == owners[REWRITE]['inclusive_ir']
                    and owners[REWRITE]['direct'][SCAN] == owners[SCAN]['inclusive_ir']
                    and owners[SCAN]['direct'][START] == owners[START]['inclusive_ir']
                    and owners[START]['direct'][ADDRESS] == owners[ADDRESS]['inclusive_ir'],
                    'selected direct owner chain differs')
            rows.append({'repeat': repeat, 'shape': shape,
                         'inputs': {str(p.relative_to(REPO)): sha(p) for p in (raw, inc_path, self_path)},
                         'owners': owners, 'raw_edges_and_self_equations_verified': True})
    for owner in totals.values():
        owner['direct'] = dict(sorted(owner['direct'].items(), key=lambda item: (-item[1], item[0])))
        owner['share_of_commit_percent'] = owner['inclusive_ir'] / totals[OWNER]['inclusive_ir'] * 100
    return {'status': 'pass', 'source_binding_sha256': sha(HERE / 'source-binding.json'),
            'prior_seal_sha256': sha(PRIOR / 'SHA256SUMS'),
            'helpers': {str(p.relative_to(REPO)): sha(p) for p in helpers},
            'rows': rows, 'aggregate': totals,
            'scope': 'Retained 0525 candidate commit Ir only. Direct children are disjoint only within their immediate parent; selected owners overlap through nesting.',
            'limitations': ['No new latency, allocation, RSS, I/O or scaling measurement.',
                            'Call metadata includes collection-off work and is not used as an allocation count.',
                            'Allocator and resolver edges are not entirely removable by the proposed storage change.']}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    text = json.dumps(analyze(), indent=2, sort_keys=True) + '\n'
    if args.output:
        args.output.write_text(text)
    else:
        print(text, end='')
