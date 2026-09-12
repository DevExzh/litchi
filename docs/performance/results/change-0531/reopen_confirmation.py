"""Capture and replay the frozen supplemental reopen regression confirmation."""
import argparse
import json
import subprocess

import analyze as numeric
from run import HERE, SCRATCH, now, plan_data, run, sha, write


def jobs():
    config = json.loads((HERE / 'reopen-plan.json').read_text())
    assert config['plan_sha256'] == sha(HERE / 'plan.json')
    for stage, repeat in config['order']:
        for index, case in enumerate(config['cases']):
            yield stage, dict(name=f'reopen-r{repeat}-guard{index}-{case["shape"]}',
                              kind='guard', guard=index, repeat=repeat,
                              samples=config['samples'], warmup=config['warmup'], **case)


def command(stage, job):
    return ['taskset', '-c', str(plan_data()['cpu']), '/usr/bin/time', '-f',
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
            '-o', str(HERE / stage / (job['name'] + '.rss.json')),
            str(SCRATCH / (stage + '-normal')), '--warmup', str(job['warmup']),
            '--samples', str(job['samples']), '--case', job['case'],
            '--xlsx-cell-crud-shape', job['shape'], '--json',
            str(HERE / stage / (job['name'] + '.json'))]


def capture():
    for stage, job in jobs():
        rows = subprocess.check_output(['ps', '-eo', 'pid,ppid,etime,comm'], text=True).splitlines()
        write(HERE / stage / (job['name'] + '.host.json'), dict(
            observed_utc=now(), processes=[r for r in rows if r.split()[-1] in ['cargo', 'rustc']],
            scope='Pre-child snapshot, not proof of an idle host.'))
        run(stage, job['name'], command(stage, job), SCRATCH / (stage + '-normal'),
            retained_baseline=stage == 'baseline')


def analyze():
    rows = {'baseline': {}, 'candidate': {}}
    receipts = []
    for stage, job in jobs():
        folder = HERE / stage
        receipt = json.loads((folder / (job['name'] + '.receipt.json')).read_text())
        assert receipt['command'] == command(stage, job)
        assert receipt['exit_code'] == 0
        assert receipt['plan_sha256'] == sha(HERE / 'plan.json')
        assert receipt['script_sha256'] == sha(HERE / 'run.py')
        assert receipt['source_manifest_sha256'] == sha(folder / 'source-manifest.json')
        assert receipt['working_source_manifest_sha256'] == sha(HERE / 'candidate/source-manifest.json')
        binary = json.loads((folder / 'binary-normal.json').read_text())
        assert receipt['binary_sha256'] == binary['sha256']
        inventory = {p.name: sha(p) for p in folder.glob(job['name'] + '.*')
                     if p.is_file() and not p.name.endswith('.receipt.json')}
        assert receipt['artifacts'] == inventory
        receipts.append(receipt)
        raw = json.loads((folder / (job['name'] + '.json')).read_text())
        row = numeric._validate_result(raw, plan_data(), job, binary, False)
        row['rss'] = numeric.BASE.validate_rss(folder / (job['name'] + '.rss.json'))
        rows[stage][(job['repeat'], job['guard'])] = row
    assert all(a['end_utc'] <= b['start_utc'] for a, b in zip(receipts, receipts[1:]))
    comparisons, adverse, drift = [], [], []
    for key, left in rows['baseline'].items():
        right = rows['candidate'][key]
        assert left['identity'] == right['identity']
        values = []
        for phase in left['timing']:
            for stat in ['p50', 'p95', 'p99', 'mean']:
                a, b = left['timing'][phase][stat], right['timing'][phase][stat]
                change = (b / a - 1) * 100 if a else (0 if b == 0 else None)
                value = dict(phase=phase, stat=stat, baseline=a, candidate=b, change_percent=change)
                values.append(value)
                if change is None or change > 5:
                    adverse.append(dict(repeat=key[0], guard=key[1], **value))
        a, b = left['rss']['max_rss_kib'], right['rss']['max_rss_kib']
        rss = dict(phase='rss', stat='max_rss_kib', baseline=a, candidate=b,
                   change_percent=(b / a - 1) * 100)
        values.append(rss)
        if rss['change_percent'] > 5:
            adverse.append(dict(repeat=key[0], guard=key[1], **rss))
        gates = [v for v in values if v['phase'] in ['elapsed_ns', 'reopen_ns']
                 and v['stat'] in ['p50', 'mean']]
        assert len(gates) == 4
        comparisons.append(dict(repeat=key[0], guard=key[1], case=left['case'], shape=left['shape'],
                                metrics=values, passed=all(v['change_percent'] is not None
                                                          and v['change_percent'] <= 5 for v in gates)))
    for stage, stage_rows in rows.items():
        for guard in range(3):
            a, b = stage_rows[(1, guard)], stage_rows[(2, guard)]
            assert a['identity'] == b['identity']
            for phase in a['timing']:
                for stat in ['p50', 'p95', 'p99', 'mean']:
                    x, y = a['timing'][phase][stat], b['timing'][phase][stat]
                    change = (y / x - 1) * 100 if x else (0 if y == 0 else None)
                    if change is None or abs(change) > 5:
                        drift.append(dict(stage=stage, guard=guard, phase=phase, stat=stat,
                                          repeat1=x, repeat2=y, change_percent=change))
            x, y = a['rss']['max_rss_kib'], b['rss']['max_rss_kib']
            change = (y / x - 1) * 100
            if abs(change) > 5:
                drift.append(dict(stage=stage, guard=guard, phase='rss', stat='max_rss_kib',
                                  repeat1=x, repeat2=y, change_percent=change))
    return dict(status='pass', guard_passed=all(r['passed'] for r in comparisons),
                plan_sha256=sha(HERE / 'reopen-plan.json'), script_sha256=sha(HERE / 'reopen_confirmation.py'),
                comparisons=comparisons, adverse=adverse, same_build_drift=drift,
                scope='Fresh supplemental reopen and edit/save regression guard; original adverse evidence retained. No speedup claim.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('action', choices=['capture', 'analyze'])
    args = parser.parse_args()
    if args.action == 'capture':
        capture()
    else:
        result = analyze()
        write(HERE / 'reopen-comparison.json', result)
        print('Reopen confirmation guard:', result['guard_passed'])
