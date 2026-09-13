"""Frozen job matrix with serial source/binary-bound capture receipts."""
import json
import run as R


def jobs(lane):
    p = json.loads((R.HERE / 'plan.json').read_text())
    config = p['native' if lane == 'preflight' else lane]
    for repeat in range(1, (1 if lane == 'preflight' else config['repeats']) + 1):
        shapes = p['shapes'] if repeat == 1 else list(reversed(p['shapes']))
        cases = [p['profile']['case']] if lane == 'profile' else p['cases']
        for shape in shapes:
            for i, case in enumerate(cases):
                yield dict(name=f'{lane}-r{repeat}-{shape}-c{i}', lane=lane,
                    repeat=repeat, shape=shape, case=case,
                    samples=1 if lane == 'preflight' else config['samples'],
                    warmup=0 if lane == 'preflight' else config['warmup'])


def capture(lane, repeat=None):
    p = json.loads((R.HERE / 'plan.json').read_text())
    kind = 'alloc' if lane == 'alloc' else 'normal'
    binary = R.SCRATCH / kind
    identity = json.loads((R.FOLDER / ('binary-' + kind + '.json')).read_text())
    assert R.sha(binary) == identity['sha256']
    for job in jobs(lane):
        if repeat is not None and job['repeat'] != repeat:
            continue
        name = job['name']
        command = ['taskset', '-c', str(p['cpu']), '/usr/bin/time', '-v']
        if lane == 'profile':
            owner = p['profile']['owner']
            command += ['valgrind', '--tool=callgrind', '--collect-atstart=no',
                '--toggle-collect=' + owner, '--zero-before=' + owner,
                '--dump-after=' + owner, '--callgrind-out-file=' + str(R.FOLDER / (name + '.callgrind'))]
        command += [str(binary), '--case', job['case'], '--xlsx-cell-crud-shape', job['shape'],
            '--warmup', str(job['warmup']), '--samples', str(job['samples']),
            '--json', str(R.FOLDER / (name + '.json')),
            '--corpus-manifest', str(R.FOLDER / (name + '.catalog.json'))]
        R.run(name, command, binary)


def main():
    (R.TARGET / 'tmp').mkdir(parents=True, exist_ok=False)
    R.freeze()
    R.build('normal')
    R.run('symbols', ['nm', '-C', str(R.SCRATCH / 'normal')], R.SCRATCH / 'normal')
    p = json.loads((R.HERE / 'plan.json').read_text())
    lines = (R.FOLDER / 'symbols.stdout').read_text().splitlines()
    selected = [line for line in lines if line.endswith(' ' + p['profile']['owner'])]
    assert len(selected) == 1, selected
    R.write(R.HERE / 'symbol-observation.json', {'owner':p['profile']['owner'], 'rows':selected,
        'plan_sha256':R.sha(R.HERE / 'plan.json'), 'symbols_sha256':R.sha(R.FOLDER / 'symbols.stdout')})
    capture('preflight')
    capture('native', 1)
    capture('profile', 1)
    capture('profile', 2)
    capture('native', 2)
    R.build('alloc')
    capture('alloc', 1)
    capture('alloc', 2)
    for name, command in [
        ('fmt', ['cargo','fmt','--all','--check']),
        ('harness-fmt', ['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--all','--check']),
        ('boundaries', ['python3','-B','tools/check_crate_boundaries.py']),
        ('claims', ['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'])]:
        R.run('check-' + name, command)


if __name__ == '__main__':
    main()
