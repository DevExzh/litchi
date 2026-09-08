#!/usr/bin/env python3
"""Execute frozen paired measurements without overwriting any prior attempt."""
from pathlib import Path
import subprocess
import sys
from common import ROOT, ENV, ENV_KEYS, read, write, sha, meta, now, check_arm

ORDER_KEYS = {'main': 'order', 'pilot': 'pilot_order', 'counter': 'counter_order', 'guard': 'guard_order'}


def capture(suite, lane):
    protocol = read(ROOT / 'protocol.json')
    assert protocol['drivers']['capture.py'] == sha(Path(__file__))
    assert protocol['drivers']['common.py'] == sha(ROOT / 'common.py')
    item = next(row for row in protocol[ORDER_KEYS[suite]] if row['lane'] == lane)
    build = check_arm(item['arm'])
    binary = Path(build['binaries'][item['mode']]['path'])
    output = ROOT / ('pilots' if suite == 'pilot' else 'captures') / lane
    output.mkdir(parents=True, exist_ok=False)
    samples = protocol['pilot_samples'] if suite == 'pilot' else protocol['samples']
    warmups = protocol['pilot_warmups'] if suite == 'pilot' else protocol['warmups']
    cases = ','.join(protocol['guard_selectors']) if suite == 'guard' else protocol['selector']
    argv = ['taskset', '-c', str(protocol['cpu']), '/usr/bin/time', '-v', '-o', str(output / 'resource.log')]
    if suite == 'counter':
        argv += ['perf', 'stat', '-x', ';', '-o', str(output / 'counters.csv'), '-e', protocol['counter_events'], '--']
    argv += [str(binary), '--workers', str(protocol['workers']), '--warmup', str(warmups), '--samples', str(samples),
        '--case', cases, '--semantic-shape', item['shape'], '--json', str(output / 'report.json'),
        '--corpus-manifest', str(output / 'corpus-catalog.json')]
    record = dict(schema='litchi-0476-capture-v1', suite=suite, **item, argv=argv, cwd=build['build_path'],
        samples=samples, warmups=warmups, case_filter=cases, revision=build['revision'],
        binary_sha256=sha(binary), build_sha256=sha(ROOT / f"{item['arm']}-build.json"),
        protocol_sha256=sha(ROOT / 'protocol.json'), driver_sha256=sha(Path(__file__)),
        common_sha256=sha(ROOT / 'common.py'), environment={k: ENV[k] for k in ENV_KEYS},
        clean_before=True, started_utc=now())
    write(output / 'started.json', record)
    with (output / 'stdout.log').open('xb') as out, (output / 'stderr.log').open('xb') as err:
        process = subprocess.run(argv, cwd=build['build_path'], env=ENV, stdout=out, stderr=err)
    check_arm(item['arm'])
    record.update(exit_code=process.returncode, finished_utc=now(), clean_after=True,
        binary_unchanged=True, source_unchanged=True,
        artifacts={p.name: meta(p) for p in sorted(output.iterdir()) if p.is_file()})
    write(output / 'receipt.json', record)
    print(suite, lane, process.returncode, flush=True)
    assert process.returncode == 0
    report = read(output / 'report.json')
    assert report['environment']['git_revision'] == build['revision']
    assert report['environment']['git_worktree_dirty'] is False
    rows = report['results']
    if suite != 'guard':
        assert len(rows) == 1
        assert rows[0]['case'] == protocol['selector']
        assert rows[0]['output_sha256'] == protocol['corpora'][item['shape']]['archive_sha256']
    else:
        assert len(rows) == len(protocol['guard_selectors']) * len(item['shape'].split(','))
        assert {r['case'] for r in rows} == set(protocol['guard_selectors'])
    for row in rows:
        assert row['output_sha256'] == row['corpus']['archive_sha256']
        assert len(row['elapsed_ns']['samples']) == samples


if __name__ == '__main__':
    suite = sys.argv[1]
    assert suite in ORDER_KEYS
    if len(sys.argv) == 3:
        capture(suite, sys.argv[2])
    else:
        for item in read(ROOT / 'protocol.json')[ORDER_KEYS[suite]]:
            capture(suite, item['lane'])
