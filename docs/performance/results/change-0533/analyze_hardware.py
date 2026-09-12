"""Validate 0533 grouped whole-child perf counters; never relabel them operation-local."""
import argparse
import csv
import json
from pathlib import Path

import analyze as numeric
from analyze import read, binary, receipt, validate_report, identity
from run import HERE, FOLDER, jobs, sha

GROUP = ('cycles', 'instructions', 'branches', 'branch-misses')
EVENTS = GROUP + ('page-faults', 'context-switches', 'cpu-migrations')


def analyze(stage):
    global FOLDER
    numeric.configure(stage)
    FOLDER = HERE / stage
    plan = read(HERE / 'plan.json')
    binary('normal')
    native = read(FOLDER / 'native-r1-xls.json')['results']
    expected_identity = identity(next(r for r in native if r['case'] == plan['hardware']['case']))
    results = []
    for job in jobs('hardware'):
        name = job['name']
        r = receipt(name, 'normal', allow_failure=True)
        command = r['command']
        assert command[:5] == ['taskset', '-c', str(plan['cpu']), 'perf', 'stat']
        assert command[command.index('-e') + 1] == plan['hardware']['events']
        result = dict(name=name, exit_code=r['exit_code'], receipt_sha256=sha(FOLDER / (name + '.receipt.json')))
        if r['exit_code'] != 0:
            result.update(status='unavailable', reason='perf child failed; raw stderr/CSV and exit code retained, no counter claim')
            results.append(result)
            continue
        rows = validate_report(FOLDER / (name + '.json'), job)
        assert len(rows) == 1 and identity(rows[0]) == expected_identity
        events = {}
        for fields in csv.reader((FOLDER / (name + '.csv')).read_text().splitlines()):
            if not fields or fields[0].startswith('#'):
                continue
            assert len(fields) >= 5, fields
            count, unit, event, runtime, running = fields[:5]
            assert event in EVENTS and event not in events, fields
            if count.startswith('<') or not count:
                events[event] = dict(status='unavailable', raw=fields)
            else:
                value, duration, fraction = int(count), int(runtime), float(running)
                assert value >= 0 and duration > 0 and 0 <= fraction <= 100
                events[event] = dict(status='measured', value=value, event_runtime_ns=duration, running_percent=fraction)
        assert set(events) == set(EVENTS)
        group = [events[event] for event in GROUP]
        valid = all(e['status'] == 'measured' and e['running_percent'] == 100 for e in group)
        valid = valid and len({e['event_runtime_ns'] for e in group}) == 1
        result.update(events=events, identity_equal=True, group_valid=valid)
        if valid:
            values = {n: events[n]['value'] for n in GROUP}
            assert values['cycles'] > 0 and values['branches'] >= values['branch-misses']
            result.update(status='measured', ipc=values['instructions']/values['cycles'],
                branch_miss_percent=100*values['branch-misses']/values['branches'])
        else:
            result.update(status='unavailable_for_group_claim', reason='Missing or multiplexed grouped event; raw counts retained without IPC/group claim')
        results.append(result)
    return dict(status='pass', captures=results, scope=plan['hardware']['scope'],
        latency_samples_excluded_from_native=plan['hardware']['samples']*len(results),
        no_operation_local_hardware_or_speedup_claim=True)


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('output',nargs='?',type=Path)
    parser.add_argument('--stage',choices=['baseline','candidate'],default='baseline')
    args=parser.parse_args()
    output=args.output or HERE/args.stage/'hardware-analysis.json'
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(analyze(args.stage),indent=2)+'\n')
    print('Hardware diagnostic verified:',output)
