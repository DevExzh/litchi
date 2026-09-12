"""Admit the fourth separately dumped commit only after raw call-scope checks."""
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys

from capture import HERE, sha, source_check

# Reuse the retained raw-edge parser, without changing earlier evidence.
HELPERS = HERE.parent / 'change-0519'
spec = importlib.util.spec_from_file_location('prior_raw_profiles', HELPERS / 'analyze_profiles.py')
raw = importlib.util.module_from_spec(spec)
spec.loader.exec_module(raw)
sys.modules['analyze_profiles'] = raw
spec = importlib.util.spec_from_file_location('prior_profile_edges', HELPERS / 'compare_profile_lanes.py')
edges = importlib.util.module_from_spec(spec)
spec.loader.exec_module(edges)


def selected_row(text, owner):
    lines = text.splitlines()
    indexes = [i for i, line in enumerate(lines)
               if (m := edges.STAR_RE.match(line)) and edges.display_name(m[2]) == owner]
    assert len(indexes) == 1, (owner, indexes)
    i = indexes[0]
    cost = int(edges.STAR_RE.match(lines[i])[1].replace(',', ''))
    direct = []
    for line in lines[i+1:]:
        m = edges.EDGE_RE.match(line)
        if not m:
            break
        calls = re.search(r'\(([\d,]+)x\)', m[2])
        direct.append(dict(owner=edges.display_name(m[2]), ir=int(m[1].replace(',', '')),
                           calls=int(calls[1].replace(',', '')) if calls else None))
    return cost, direct


def analyze():
    source_check()
    plan = json.loads((HERE / 'profile-plan.json').read_text())
    native = json.loads((HERE / 'analysis.json').read_text())
    build = json.loads((HERE / 'build.json').read_text())
    owner = plan['owner']
    rows = []
    for repeat in range(1, plan['repeats'] + 1):
        for shape in plan['shapes']:
            name = f'profile-r{repeat}-{shape}'
            receipt = json.loads((HERE / (name + '.receipt.json')).read_text())
            assert receipt['exit_code'] == 0
            assert receipt['binary_sha256'] == build['binary_sha256']
            for filename, digest in receipt['artifacts'].items():
                assert sha(HERE / filename) == digest, filename
            assert all(option in receipt['command'] for option in plan['options'])
            dumps = sorted(HERE.glob(name + '.callgrind.[0-9]*'))
            assert [p.suffix for p in dumps] == ['.1', '.2', '.3', '.4']
            scopes = []
            for index, path in enumerate(dumps, 1):
                text = path.read_text()
                assert f'part: {index}\n' in text
                assert f'desc: Trigger: --dump-after={owner}\n' in text
                summary = int(re.search(r'^summary: (\d+)$', text, re.M)[1])
                incoming = edges.raw_incoming_call_summary(path, owner)
                assert incoming['calls'] == 1 and incoming['positive_edge_count'] == 1
                edge = incoming['edges'][0]
                parent = 'run_xlsx_cell_values_edit_save' if index == 4 else 'run_xlsx_cell_value_lifecycle_gates'
                assert edge['caller'] == 'litchi_perf_baseline::' + parent
                assert edge['inclusive_ir'] == summary
                scopes.append(dict(file=path.name, sha256=sha(path), summary_ir=summary, incoming=edge))
            raw_result = json.loads((HERE / (name + '.json')).read_text())
            assert raw_result['binary_identity']['binary_sha256'] == build['binary_sha256']
            assert raw_result['configuration']['warmup_iterations_per_case'] == 0
            assert raw_result['configuration']['samples_per_case'] == 1
            result = raw_result['results'][0]
            reference = native['identities'][shape]
            assert result['corpus'] == reference['corpus']
            assert result['sink'] == reference['sink']
            values = result['source']['xlsx_cell_values']
            for key, value in reference['constants'].items():
                actual = values[key]
                assert (actual[0] if isinstance(actual, list) else actual) == value, key
            texts = {}
            for kind in ['inclusive', 'self']:
                command = ['callgrind_annotate', '--auto=no', '--threshold=100', '--show-percs=no',
                           '--inclusive=' + ('yes' if kind == 'inclusive' else 'no'), '--tree=both', str(dumps[3])]
                # callgrind_annotate is Perl; stabilize equal-cost row order.
                environment = dict(os.environ, PERL_HASH_SEED='0', PERL_PERTURB_KEYS='0')
                process = subprocess.run(command, capture_output=True, text=True, check=True, env=environment)
                assert not process.stderr
                texts[kind] = process.stdout
                (HERE / (name + '.' + kind + '.txt')).write_text(process.stdout)
            inclusive, direct = selected_row(texts['inclusive'], owner)
            self_cost, self_direct = selected_row(texts['self'], owner)
            assert direct == self_direct
            assert inclusive == scopes[3]['summary_ir'] == self_cost + sum(x['ir'] for x in direct)
            rows.append(dict(name=name, shape=shape, scopes=scopes, selected_inclusive_ir=inclusive,
                             selected_self_ir=self_cost, direct_callees=direct))
    return dict(status='pass', scope='one timed MultiSourceEdit::commit per profile; staged set loop excluded',
                rows=rows, helpers={p.name: sha(p) for p in [HELPERS/'analyze_profiles.py', HELPERS/'compare_profile_lanes.py']},
                limitation='Valgrind reports brk segment overflow notices in retained stderr. All children complete with matching native correctness/counter evidence. Guest instruction counts are not hardware instructions, cycles, or native timings.')


if __name__ == '__main__':
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / 'profile-analysis.json'
    output.write_text(json.dumps(analyze(), indent=2) + '\n')
    print('Profile call scopes and raw accounting verified:', output)
