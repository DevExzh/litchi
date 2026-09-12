"""Serial, source-bound captures of the unchanged managed DOCX example."""
import argparse
import csv
import datetime
import importlib.util
import json
import subprocess
import time

from run import HERE, REPO, SCRATCH, sha, sources, write

spec = importlib.util.spec_from_file_location(
    'docx0500', HERE.parent / 'change-0500/verify-evidence.py')
prior = importlib.util.module_from_spec(spec)
spec.loader.exec_module(prior)


def validate(name, rows, samples, warmups, repeats):
    p, k, provider, mode = prior.parse_name(name)
    expected = [(repeat, warm, ordinal) for repeat in range(repeats)
                for warm, count in [('true', warmups), ('false', samples)]
                for ordinal in range(count)]
    assert [(int(r['repeat']), r['warmup'], int(r['ordinal'])) for r in rows] == expected
    assert len({tuple(r[key] for key in prior.IDENTITY_KEYS) for r in rows}) == 1
    for r in rows:
        assert r['schema'] == 'managed_paragraph_batch_perf_v1' and r['version'] == '1'
        assert (int(r['paragraphs']), int(r['replacements']), r['source'], r['mode']) == (p, k, provider, mode)
        assert r['api_path'] == ('replace_paragraph_text' if mode == 'repeated' else 'replace_body_paragraph_texts')
        assert all(r[key] == 'true' for key in prior.BOOLEAN_KEYS)
        assert r['budget_managed'] == 'true'
        assert r['expected_output_sha256'] == r['output_sha256']
        assert int(r['expected_output_bytes']) == int(r['output_bytes']) > 0
        assert all(int(r[key]) >= 0 for key in prior.TIMED_FIELDS)
        assert int(r['elapsed_ns']) >= sum(int(r[key]) for key in prior.TIMED_FIELDS if key != 'elapsed_ns')
        for budget in ['memory', 'objects']:
            assert int(r['budget_before_' + budget]) == int(r['budget_after_' + budget]) == 0
        for budget in ['input', 'work']:
            assert int(r['budget_after_' + budget]) >= int(r['budget_live_' + budget]) >= int(r['budget_before_' + budget])
        assert (r['source_version_before_id'], r['source_version_before_revision']) == (r['source_version_after_id'], r['source_version_after_revision'])
    if (samples, warmups, repeats) == (30, 3, 2):
        prior.validate_data_rows(name, rows)


def capture(lane, name):
    plan = json.loads((HERE / 'plan.json').read_text())
    variant = 'candidate' if 'after-' in lane or lane == 'hardware-after' else 'baseline'
    build = json.loads((HERE / variant / 'build-receipt.json').read_text())
    binary = SCRATCH / ('managed-paragraph-' + variant)
    assert sha(binary) == build['binary_sha256']
    current = sources()
    assert current == json.loads((HERE / variant / 'source-manifest.json').read_text())
    directory = HERE / lane
    directory.mkdir(exist_ok=True)
    profile = lane.startswith('profile')
    hardware = lane in ['hardware', 'hardware-after']
    small = lane == 'preflight' or profile
    samples, warmups, repeats = (1, 0, 1) if small else (30, 3, 2)
    p, k, provider, mode = prior.parse_name(name)
    output = directory / (name + '.csv')
    stdout = directory / (name + '.stdout')
    stderr = directory / (name + '.stderr')
    receipt = directory / (name + '.json')
    scratch = SCRATCH / 'corpora' / (lane + '-' + name)
    assert not scratch.exists()
    scratch.parent.mkdir(exist_ok=True)
    for path in [output, stdout, stderr, receipt]:
        assert not path.exists(), path
    command = [str(binary), '--paragraphs', str(p), '--replacements', str(k),
               '--source', provider, '--mode', mode, '--samples', str(samples),
               '--warmups', str(warmups), '--repeats', str(repeats),
               '--artifact-dir', str(scratch), '--output', str(output)]
    artifacts = [output, stdout, stderr]
    if profile:
        raw = directory / (name + '.callgrind')
        assert not raw.exists()
        command = ['valgrind', '--tool=callgrind', '--collect-atstart=no',
                   '--zero-before=' + plan['profile']['zero_before'],
                   '--toggle-collect=' + plan['profile']['toggle_collect'],
                   '--callgrind-out-file=' + str(raw), *command]
        artifacts.append(raw)
    elif hardware:
        counters = directory / (name + '.perf.csv')
        assert not counters.exists()
        command = ['perf', 'stat', '-x,', '-o', str(counters), '-e',
                   ','.join(plan['hardware']['events']), '--', *command]
        artifacts.append(counters)
    command = ['/usr/bin/time', '-v', 'taskset', '-c', '2', *command]
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with stdout.open('x') as out, stderr.open('x') as err:
        result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
    unchanged = sources() == current
    record = {'command': command, 'started_utc': start,
              'elapsed_seconds': time.monotonic() - tick,
              'exit_code': result.returncode, 'source_unchanged': unchanged,
              'source_manifest_sha256': sha(HERE / variant / 'source-manifest.json'),
              'binary_sha256': sha(binary), 'plan_sha256': sha(HERE / 'plan.json'),
              'samples': samples, 'warmups': warmups, 'repeats': repeats,
              'cleanup_verified': not scratch.exists(),
              'scope': plan['profile']['scope'] if profile else plan['hardware']['scope'] if hardware else plan['native']['scope'],
              'artifacts': {p.name: sha(p) for p in artifacts if p.exists()}}
    if variant == 'candidate':
        record['candidate_plan_sha256'] = sha(HERE / 'candidate-plan.json')
    write(receipt, record)
    assert result.returncode == 0 and unchanged and not scratch.exists(), name
    with output.open(newline='') as stream:
        rows = list(csv.DictReader(stream))
    validate(name, rows, samples, warmups, repeats)
    print(lane, name, len(rows), 'rows passed', flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('lane', choices=['preflight', 'r1', 'r2', 'profile-preflight', 'profile-r1', 'profile-r2', 'hardware', 'after-r1', 'after-r2', 'profile-after-r1', 'profile-after-r2', 'hardware-after'])
    parser.add_argument('--case')
    args = parser.parse_args()
    plan = json.loads((HERE / 'plan.json').read_text())
    if args.case:
        names = [args.case]
    elif args.lane.startswith('profile') or args.lane in ['hardware', 'hardware-after']:
        names = ['p128-k1-owned-batch'] if args.lane == 'profile-preflight' else plan['profile']['cases']
    else:
        names = [f'p{p}-k{k}-{provider}-{mode}' for p in [128, 512]
                 for k in [1, 8, 32] for provider in ['owned', 'file']
                 for mode in ['repeated', 'batch']]
    if args.lane in ['r1', 'r2', 'after-r1', 'after-r2']:
        names.sort()
    if args.lane.endswith('r2'):
        names.reverse()
    for name in names:
        capture(args.lane, name)


if __name__ == '__main__':
    main()
