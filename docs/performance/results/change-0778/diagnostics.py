"""Root-only serial syscall and marginal instruction observations, after native."""
from pathlib import Path
import os
import subprocess
import sys
import time
import custody as c


def fixtures(plan):
    for corpus in plan['corpora']:
        if corpus['path']:
            f = c.ROOT / corpus['path']
            assert f.stat().st_size == corpus['bytes'] and c.sha(f) == corpus['sha256']


def arguments(plan, binary, corpus, policy, samples, warmup, report):
    prefix = corpus['format'] + ('_real_file' if corpus['path'] else '')
    args = [binary['path'], '--case', prefix + '_ordinary_save_atomic_publish',
            '--samples', str(samples), '--warmup', str(warmup),
            '--filesystem-root', plan['filesystem_root'], '--json', str(report)]
    if corpus['path']:
        args += ['--ooxml-file', corpus['path']]
    if policy != 'default':
        args += ['--save-durability', policy]
    return args


def execute(command, prefix, fields):
    stdout, stderr = prefix.with_suffix('.stdout'), prefix.with_suffix('.stderr')
    started = time.time()
    with stdout.open('wb') as out, stderr.open('wb') as err:
        r = subprocess.run(command, cwd=c.ROOT, env=os.environ | {'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC'}, stdout=out, stderr=err)
    return fields | {'command': command, 'exit_code': r.returncode,
                     'started': started, 'ended': time.time(),
                     'stdout': c.artifact(stdout), 'stderr': c.artifact(stderr)}


if __name__ == '__main__':
    lane = sys.argv[1]
    assert lane in ['trace', 'instructions']
    plan, build = c.read(c.P / 'plan.json'), c.read(c.P / 'build.json')
    source = c.read(c.P / 'source.json')
    assert c.census() == source
    assert c.read(c.P / 'admission.json')['oracle_pass'] is True
    # The caller runs this only after capture.py's native lane is terminal.
    native_complete = c.P / 'native-0/complete.json'
    assert native_complete.is_file(), 'native lane must complete first'
    binary = build['binaries']['native']
    assert c.artifact(Path(binary['path'])) == binary
    fixtures(plan)
    out = c.P / f'{lane}-0'
    assert not out.exists()
    out.mkdir()
    script_sha = c.sha(Path(__file__))
    rows = []
    if lane == 'trace':
        settings = plan['diagnostics']['trace']
        for corpus in plan['corpora']:
            for policy in plan['policies']:
                prefix = out / f'{corpus["id"]}-{policy}'
                report, trace = prefix.with_suffix('.json'), prefix.with_suffix('.strace')
                args = arguments(plan, binary, corpus, policy, settings['samples'], settings['warmup'], report)
                command = ['taskset', '-c', str(plan['cpu']), '/usr/bin/strace', *settings['options'], '-o', str(trace), *args]
                row = execute(command, prefix, {'corpus': corpus['id'], 'policy': policy, 'binary': binary})
                for field, path in [('report', report), ('trace', trace)]:
                    row[field] = c.artifact(path) if path.exists() else None
                rows.append(row)
                c.write(out / 'runs.json', rows)
                assert row['exit_code'] == 0, row
                assert c.census() == source
                fixtures(plan)
                print(prefix.name + ' traced', flush=True)
    else:
        settings = plan['diagnostics']['instructions']
        csv = out / 'qualification.csv'
        command = ['perf', 'stat', '-x', ';', '--no-big-num', '-e', 'instructions:u', '-o', str(csv), '--', 'taskset', '-c', str(plan['cpu']), '/usr/bin/true']
        row = execute(command, out / 'qualification', {})
        row['csv'] = c.artifact(csv) if csv.exists() else None
        c.write(out / 'qualification.json', row)
        supported = row['exit_code'] == 0 and csv.exists() and any(
            len(parts := line.split(';')) > 2 and parts[2] == 'instructions:u' and parts[0].strip().isdigit()
            for line in csv.read_text().splitlines())
        if not supported:
            c.write(out / 'complete.json', {'supported': False, 'reason': 'Hardware counter qualification failed; no instruction result claimed.'})
            raise SystemExit(0)
        for corpus in plan['corpora']:
            if corpus['id'] not in settings['corpora']:
                continue
            for repeat in range(settings['repeats']):
                schedule = [('default', 3), ('no-sync', 3), ('no-sync', 23), ('default', 23)]
                if repeat % 2:
                    schedule = [('no-sync', 3), ('default', 3), ('default', 23), ('no-sync', 23)]
                for policy, samples in schedule:
                    prefix = out / f'{corpus["id"]}-{repeat}-{policy}-{samples}'
                    report, csv = prefix.with_suffix('.json'), prefix.with_suffix('.csv')
                    args = arguments(plan, binary, corpus, policy, samples, settings['warmup'], report)
                    command = ['perf', 'stat', '-x', ';', '--no-big-num', '-e', ','.join(settings['events']), '-o', str(csv), '--', 'taskset', '-c', str(plan['cpu']), *args]
                    row = execute(command, prefix, {'corpus': corpus['id'], 'policy': policy, 'repeat': repeat, 'samples': samples, 'warmup': settings['warmup'], 'binary': binary})
                    for field, path in [('report', report), ('csv', csv)]:
                        row[field] = c.artifact(path) if path.exists() else None
                    rows.append(row)
                    c.write(out / 'runs.json', rows)
                    assert row['exit_code'] == 0, row
                    assert c.census() == source
                    fixtures(plan)
                print(prefix.name + ' counters captured', flush=True)
    assert len(rows) == settings['processes']
    assert c.census() == source
    assert c.artifact(Path(binary['path'])) == binary
    fixtures(plan)
    assert c.sha(Path(__file__)) == script_sha
    c.write(out / 'complete.json', {'supported': True, 'runs': len(rows), 'source_unchanged': True, 'fixtures_unchanged': True, 'binary_unchanged': True, 'runner_sha256': script_sha, 'plan_sha256': c.sha(c.P / 'plan.json'), 'source_sha256': c.sha(c.P / 'source.json')})
