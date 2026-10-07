"""Independent offline recomputation of the retained 0948 gates; no workloads."""
from __future__ import annotations
import hashlib, json, math, random, statistics
from pathlib import Path

P = Path('/home/zhuhe/litchi-performance-0948')
OUT = P / 'review/protocol/review_protocol_0948_v9-final-gates.json'
CASES = ('primary-percent-medium', 'primary-percent-dense-sparse',
         'control-xlsx-one-edit-medium', 'control-xlsx-managed-one-edit-medium')
ROLES = {'baseline', 'candidate'}
ALLOC_FIELDS = ('allocated_bytes', 'allocation_calls', 'region_peak_live_bytes',
                'live_bytes_after', 'max_rss_kib')

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def boot(values):
    rng = random.Random(948009)
    draws = [statistics.median([values[rng.randrange(len(values))] for _ in values])
             for _ in range(10000)]
    draws.sort()
    return {'estimate': float(statistics.median(values)),
            'lower_95': draws[24], 'upper_95': draws[9749],
            'replicates': 10000, 'seed': 948009}

def pair(rows, index, case):
    found = [row for row in rows if row['index'] == index and row['case_id'] == case]
    assert len(found) == 2 and {row['role'] for row in found} == ROLES
    return (next(row for row in found if row['role'] == 'baseline'),
            next(row for row in found if row['role'] == 'candidate'))

def med_metric(row, field):
    if field == 'max_rss_kib':
        return row[field]
    values = [item[field] for item in row['allocator_metrics']['publication_allocation_metrics']]
    assert len(values) == 5 and all(isinstance(value, int) and value >= 0 for value in values)
    return sorted(values)[len(values) // 2]

def close(a, b):
    return a == b or math.isclose(a, b, rel_tol=1e-12, abs_tol=1e-12)

def main():
    assert not OUT.exists(), f'refusing overwrite: {OUT}'
    plan_path = P / 'plan-v3.json'
    analysis_path = P / 'analysis.json'
    plan = json.loads(plan_path.read_text())
    analysis = json.loads(analysis_path.read_text())
    assert analysis['schema'] == 'litchi.paired-audit-reduction-0948.v1'
    assert analysis['stage'] == 'all'
    assert len(analysis['qualification']) == 16
    assert len(analysis['native']) == 48
    assert len(analysis['allocator']) == 24
    for stage, children, samples, warmup in (('qualification', 16, 2, 0),
                                              ('native', 48, 100, 10),
                                              ('allocator', 24, 5, 2)):
        capture = json.loads((P / stage / 'capture.json').read_text())
        assert capture['status'] == 'pass'
        assert capture['children'] == children
        assert capture['samples_per_child'] == samples
        assert capture['warmup_per_child'] == warmup
        assert capture['total_samples'] == children * samples
        assert len(capture['receipts']) == children
    # All outer capture terminals were recorded and successful; this audit does
    # not infer retry absence from timing data.
    terminals = {}
    for stage in ('qualification', 'native', 'allocator'):
        receipt = json.loads((P / 'receipts/root-' + stage + '-terminal-v1.json').read_text())
        assert receipt['exit_code'] == 0
        terminals[stage] = {'path': f'receipts/root-{stage}-terminal-v1.json',
                            'sha256': sha(P / f'receipts/root-{stage}-terminal-v1.json'),
                            'exit_code': receipt['exit_code'], 'session_id': receipt.get('session_id')}

    native_rows = analysis['native']
    native_cases = {}
    native_match = True
    for case in CASES:
        blocks = []
        publication = []
        workflow = []
        for block in range(6):
            base, cand = pair(native_rows, block, case)
            metrics = {
                'publication_p50': cand['publication_p50'] / base['publication_p50'],
                'whole_workflow_p50': cand['whole_workflow_p50'] / base['whole_workflow_p50'],
                'publication_p95': cand['publication_p95'] / base['publication_p95'],
                'publication_p99': cand['publication_p99'] / base['publication_p99'],
                'whole_workflow_p95': cand['p95'] / base['p95'],
                'whole_workflow_p99': cand['p99'] / base['p99'],
            }
            publication.append(metrics['publication_p50'])
            workflow.append(metrics['whole_workflow_p50'])
            blocks.append({'block': block, 'ratios': metrics,
                           'per_block_safe': all(value <= 1.05 for value in metrics.values())})
            reduced = analysis['reduction']['cases'][case]
            stored = next(item for item in reduced['publication_p50']['pairs'] if item['block'] == block)
            for metric, key in (('publication_p50', 'publication_ratio'),
                                ('whole_workflow_p50', 'whole_workflow_ratio'),
                                ('publication_p95', 'publication_p95_ratio'),
                                ('publication_p99', 'publication_p99_ratio'),
                                ('whole_workflow_p95', 'whole_workflow_p95_ratio'),
                                ('whole_workflow_p99', 'whole_workflow_p99_ratio')):
                native_match &= close(stored[key], metrics[metric])
        pub_boot = boot(publication)
        whole_boot = boot(workflow)
        reduced = analysis['reduction']['cases'][case]
        for metric, expected in (('publication_p50', pub_boot), ('whole_workflow_p50', whole_boot)):
            stored = reduced[metric]
            native_match &= all(close(stored[key], expected[key]) for key in ('estimate', 'lower_95', 'upper_95'))
        native_cases[case] = {
            'publication_p50_blocks': publication,
            'whole_workflow_p50_blocks': workflow,
            'publication_p50_bootstrap': pub_boot,
            'whole_workflow_p50_bootstrap': whole_boot,
            'max_tail_ratios': {metric: max(block['ratios'][metric] for block in blocks)
                                for metric in ('publication_p95', 'publication_p99',
                                               'whole_workflow_p95', 'whole_workflow_p99')},
            'blocks': blocks,
            'all_blocks_safe': all(block['per_block_safe'] for block in blocks),
        }

    alloc_rows = analysis['allocator']
    alloc_cases = {}
    alloc_match = True
    for case in CASES:
        repeats = []
        for repeat in range(3):
            base, cand = pair(alloc_rows, repeat, case)
            ratios = {field: med_metric(cand, field) / med_metric(base, field)
                      for field in ALLOC_FIELDS}
            stored_pair = next(item for item in analysis['allocation_reduction']['cases'][case]['pairs']
                               if item['repeat'] == repeat)
            stored_metrics = stored_pair['metrics']
            for field in ALLOC_FIELDS:
                alloc_match &= close(stored_metrics[field]['ratio'], ratios[field])
            repeats.append({'repeat': repeat, 'ratios': ratios,
                            'per_pair_safe': all(value <= 1.05 for value in ratios.values())})
        metric_summary = {}
        for field in ALLOC_FIELDS:
            values = [row['ratios'][field] for row in repeats]
            metric_summary[field] = {'pair_ratios': values,
                                     'paired_median': statistics.median(values),
                                     'max_pair_ratio': max(values),
                                     'pass': max(values) <= 1.05}
        alloc_cases[case] = {'repeats': repeats, 'metrics': metric_summary,
                             'all_pairs_safe': all(row['per_pair_safe'] for row in repeats)}

    failures = []
    for case, values in alloc_cases.items():
        for field, summary in values['metrics'].items():
            if not summary['pass']:
                failures.append({'case': case, 'metric': field,
                                 'pair_ratios': summary['pair_ratios'],
                                 'max_pair_ratio': summary['max_pair_ratio']})
    native_gate = analysis['reduction']['all_case_latency_gate_pass']
    allocator_pair_gate = analysis['allocation_reduction']['all_pair_gate_pass']
    allocator_safety_gate = analysis['allocation_reduction']['all_row_safety_gate_pass']
    assert native_match and alloc_match
    assert native_gate is True
    assert allocator_pair_gate is False and allocator_safety_gate is False
    assert analysis['status'] == 'fail'
    result = {
        'schema': 'litchi.performance.0948.review.final-gates-v1',
        'status': 'pass',
        'scope': 'independent offline recomputation from retained analysis rows; no workloads',
        'inputs': {'plan': {'path': 'plan-v3.json', 'bytes': plan_path.stat().st_size, 'sha256': sha(plan_path)},
                   'analysis': {'path': 'analysis.json', 'bytes': analysis_path.stat().st_size, 'sha256': sha(analysis_path)},
                   'reducer_status': analysis['status'], 'stage': analysis['stage']},
        'counts': {'qualification_children': 16, 'qualification_samples': 32,
                   'native_children': 48, 'native_samples': 4800,
                   'allocator_children': 24, 'allocator_samples': 120,
                   'formal_retries': 0},
        'native': {'gate_pass': native_gate, 'recomputed_rows_match': native_match,
                   'cases': native_cases},
        'allocator': {'pair_gate_pass': allocator_pair_gate,
                      'safety_gate_pass': allocator_safety_gate,
                      'recomputed_rows_match': alloc_match,
                      'cases': alloc_cases, 'failed_metrics': failures},
        'decision': {'analysis_status': analysis['status'], 'candidate_adoptable': False,
                     'reason': 'native gate passes, but allocator pair and row safety gates fail',
                     'failed_allocator_metrics': failures},
        'outer_terminals': terminals,
        'no_recapture_or_gate_change': True,
    }
    with OUT.open('x') as handle:
        handle.write(json.dumps(result, indent=2, sort_keys=True) + '\n')
        handle.flush()
    print(json.dumps({'status': result['status'], 'native_gate': native_gate,
                      'allocator_pair_gate': allocator_pair_gate,
                      'allocator_safety_gate': allocator_safety_gate,
                      'failed_metrics': len(failures), 'output': str(OUT)}))

if __name__ == '__main__':
    main()
