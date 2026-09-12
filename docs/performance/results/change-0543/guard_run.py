"""Source-bound standalone planning guard builds and ABBA captures."""
import json
import shutil
import run


def build(stage, kind):
    run.SCRATCH.mkdir(parents=True, exist_ok=True)
    command = ['env', 'CARGO_BUILD_JOBS=2', 'CARGO_INCREMENTAL=0',
               'cargo', 'build', '--release', '--locked', '--manifest-path',
               'tools/perf-baseline/Cargo.toml', '--bin', 'xlsx_planning_guard',
               '--target-dir', str(run.TARGET)]
    if kind == 'alloc':
        command += ['--features', 'allocator-metrics']
    name = 'build-guard-' + kind
    run.run(stage, name, command)
    source = run.TARGET / 'release/xlsx_planning_guard'
    binary = run.SCRATCH / (stage + '-guard-' + kind)
    shutil.copy2(source, binary)
    assert run.sha(source) == run.sha(binary)
    run.write(run.HERE / stage / ('binary-guard-' + kind + '.json'), {
        'path': str(binary), 'sha256': run.sha(binary),
        'bytes': binary.stat().st_size,
        'build_receipt_sha256': run.sha(run.HERE / stage / (name + '.receipt.json')),
        'source_manifest_sha256': run.sha(run.HERE / stage / 'source-manifest.json'),
    })


def capture(stage, lane, repeat):
    plan = run.plan_data()
    config = plan['refusal_guard']
    kind = 'alloc' if lane == 'alloc' else 'normal'
    identity = json.loads((run.HERE / stage / ('binary-guard-' + kind + '.json')).read_text())
    binary = run.SCRATCH / (stage + '-guard-' + kind)
    assert run.sha(binary) == identity['sha256']
    prefix = 'allocation' if lane == 'alloc' else 'native'
    for shape in config['shapes'][::1 if repeat == 1 else -1]:
        for case in config['cases']:
            name = f'guard-{lane}-r{repeat}-{shape}-{case}'
            command = ['taskset', '-c', str(plan['cpu']), str(binary),
                       '--shape', shape, '--case', case,
                       '--warmup', str(config[prefix + '_warmup']),
                       '--samples', str(config[prefix + '_samples']),
                       '--json', str(run.HERE / stage / (name + '.json'))]
            run.run(stage, name, command, binary, retained_baseline=(stage == 'baseline'))
