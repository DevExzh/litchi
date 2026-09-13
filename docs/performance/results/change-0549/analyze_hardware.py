"""Validate 0549 grouped whole-child perf counters.

Perf counters are retained as whole-child diagnostics.  They never become an
operation-local instruction, cycle, or speedup claim, and unavailable or
multiplexed groups remain explicitly unavailable.
"""
import argparse
import csv
import json
import math
from pathlib import Path

import analyze as numeric
from analyze import read, binary, receipt, validate_report, identity
from run import HERE, FOLDER, jobs, sha

GROUP = ('cycles', 'instructions', 'branches', 'branch-misses')
EVENTS = GROUP + ('page-faults', 'context-switches', 'cpu-migrations')


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory hardware artifact."""


def require(condition, message):
    if not condition:
        raise EvidenceError(message)


def write_report(output: Path, value: dict) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + '\n').encode()
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f'refusing to read non-regular report output {output}')
        require(output.read_bytes() == data,
                f'refusing to overwrite non-identical report {output}')
        return
    require(not (HERE / 'SHA256SUMS').exists()
            or not output.resolve().is_relative_to(HERE.resolve()),
            f'sealed evidence is missing report output {output}')
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)


def analyze(stage):
    global FOLDER
    require(stage in ('baseline', 'candidate'), f'unsupported stage {stage!r}')
    numeric.configure(stage)
    FOLDER = HERE / stage
    plan = read(HERE / 'plan.json')
    require(plan.get('scope') ==
            'Matched CFB checked bitset test-and-mark experiment',
            'plan scope differs from the 0549 hardware lane')
    metadata = binary('normal')
    native_report = read(FOLDER / 'native-r1-xls.json')
    native = native_report.get('results')
    require(isinstance(native, list), 'native-r1-xls results are not a list')
    matching = [row for row in native
                if isinstance(row, dict)
                and row.get('case') == plan['hardware']['case']]
    require(len(matching) == 1,
            'hardware identity case is not unique in native r1 XLS')
    expected_identity = identity(matching[0])
    results = []
    for job in jobs('hardware'):
        name = job['name']
        r = receipt(name, 'normal', allow_failure=True)
        command = r['command']
        require(command[:5] == ['taskset', '-c', str(plan['cpu']), 'perf', 'stat'],
                f'{name} perf command prefix differs')
        require(command.count('-e') == 1 and
                command[command.index('-e') + 1] == plan['hardware']['events'],
                f'{name} perf events differ')
        result = dict(name=name, repeat=job['repeat'], exit_code=r['exit_code'],
                      receipt_sha256=sha(FOLDER / (name + '.receipt.json')))
        if r['exit_code'] != 0:
            result.update(status='unavailable', reason='perf child failed; raw stderr/CSV and exit code retained, no counter claim')
            results.append(result)
            continue
        rows = validate_report(FOLDER / (name + '.json'), job)
        require(len(rows) == 1 and identity(rows[0]) == expected_identity,
                f'{name} identity differs from native r1 XLS')
        events = {}
        for fields in csv.reader((FOLDER / (name + '.csv')).read_text().splitlines()):
            if not fields or fields[0].startswith('#'):
                continue
            require(len(fields) >= 5, f'{name} malformed perf row: {fields!r}')
            count, unit, event, runtime, running = fields[:5]
            require(event in EVENTS and event not in events,
                    f'{name} unexpected or duplicate event row: {fields!r}')
            if count.startswith('<') or not count:
                events[event] = dict(status='unavailable', raw=fields)
            else:
                try:
                    value, duration, fraction = int(count), int(runtime), float(running)
                except ValueError as error:
                    raise EvidenceError(f'{name} malformed numeric perf row: {fields!r}') from error
                require(value >= 0 and duration > 0 and math.isfinite(fraction)
                        and 0 <= fraction <= 100,
                        f'{name} invalid perf values: {fields!r}')
                events[event] = dict(status='measured', value=value, event_runtime_ns=duration, running_percent=fraction)
        require(set(events) == set(EVENTS),
                f'{name} event set differs: {sorted(events)}')
        group = [events[event] for event in GROUP]
        valid = all(e['status'] == 'measured' and e['running_percent'] == 100 for e in group)
        valid = valid and len({e['event_runtime_ns'] for e in group}) == 1
        result.update(events=events, identity_equal=True, group_valid=valid)
        if valid:
            values = {n: events[n]['value'] for n in GROUP}
            require(values['cycles'] > 0 and values['branches'] > 0
                    and values['branches'] >= values['branch-misses'],
                    f'{name} grouped perf counters do not reconcile')
            result.update(status='measured', ipc=values['instructions']/values['cycles'],
                branch_miss_percent=100*values['branch-misses']/values['branches'])
        else:
            result.update(status='unavailable_for_group_claim', reason='Missing or multiplexed grouped event; raw counts retained without IPC/group claim')
        results.append(result)
    return dict(schema='cfb_ole2_hardware_analysis_v1', status='pass', stage=stage,
        scope=plan['hardware']['scope'], plan_sha256=sha(HERE / 'plan.json'),
        binary_sha256=metadata['sha256'], captures=results,
        latency_samples_excluded_from_native=plan['hardware']['samples']*len(results),
        no_operation_local_hardware_or_speedup_claim=True,
        limitations=['Whole-child counters include setup, copies, queries, oracles, drop, and report construction.'])


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('output', nargs='?', type=Path,
        help='JSON destination (defaults to the stage hardware analysis)')
    parser.add_argument('--output', dest='output_option', type=Path,
        help='JSON destination (alternative to positional path)')
    parser.add_argument('--stage',choices=['baseline','candidate'],default='baseline')
    args=parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error('provide output either positionally or with --output')
    output=args.output_option or args.output or HERE/args.stage/'hardware-analysis.json'
    try:
        result = analyze(args.stage)
        write_report(output, result)
    except (EvidenceError, AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        parser.exit(2, f'analyze_hardware.py: evidence check failed: {error}\n')
    print('Hardware diagnostic verified:', output)
