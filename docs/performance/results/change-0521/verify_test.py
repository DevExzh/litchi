"""Exercise raw semantic gates independently of artifact checksum rejection."""
import copy
import json

import analyze
from run import HERE


def main():
    plan = json.loads((HERE / 'plan.json').read_text())
    checks = []
    for case in ('phase_alignment', 'generic_counter', 'output_identity', 'allocation_balance'):
        allocator = case == 'allocation_balance'
        job = analyze.expected_jobs(plan, 'allocation' if allocator else 'native')[0]
        folder = HERE / 'baseline'
        original = json.loads((folder / (job['name'] + '.json')).read_text())
        binary = json.loads((folder / ('binary-alloc.json' if allocator else 'binary-normal.json')).read_text())
        analyze.validate_result(original, plan, job, binary, allocator)
        raw = copy.deepcopy(original)
        result = raw['results'][0]
        source = result['source']['xlsx_cell_values']
        if case == 'phase_alignment':
            source['commit_ns'][0] += 1
        elif case == 'generic_counter':
            result['source']['read_calls'][0] += 1
        elif case == 'output_identity':
            result['output_sha256'] = '0' * 64
        else:
            source['commit_allocation_metrics'][0]['allocated_bytes'] += 1
        try:
            analyze.validate_result(raw, plan, job, binary, allocator)
        except analyze.EvidenceError as error:
            checks.append(dict(case=case, rejected=True, reason=str(error)))
        else:
            raise AssertionError('Corrupt evidence admitted: ' + case)
        analyze.validate_result(original, plan, job, binary, allocator)
    report = dict(status='pass', checks=checks, valid_controls_pass=True,
                  scope='In-memory raw semantic mutations; retained artifacts are unchanged.')
    (HERE / 'verifier-tests.json').write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report))


if __name__ == '__main__':
    main()
