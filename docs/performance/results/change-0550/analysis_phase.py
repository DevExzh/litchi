"""Run canonical analyses serially and retain source-bound execution receipts."""
import datetime
import json
import subprocess
import sys
import run as R


def execute(name, script, output):
    for path, digest in json.loads((R.HERE / 'analysis-inputs.json').read_text()).items():
        assert R.sha(R.REPO / path) == digest, path
    folder = R.HERE / 'analysis-runs' / name
    folder.mkdir(parents=True, exist_ok=False)
    destination = R.HERE / output
    assert not destination.exists()
    command = [sys.executable, '-B', str(R.HERE / script), '--output', str(destination)]
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with (folder / 'stdout').open('x') as out, (folder / 'stderr').open('x') as err:
        child = subprocess.run(command, cwd=R.REPO, stdout=out, stderr=err)
    receipt = dict(command=command, start_utc=start, end_utc=R.now(),
        exit_code=child.returncode, script_sha256=R.sha(R.HERE / script),
        plan_sha256=R.sha(R.HERE / 'plan.json'), output=output,
        output_sha256=R.sha(destination) if destination.exists() else None,
        artifacts={p.name:R.sha(p) for p in folder.iterdir() if p.is_file()},
        inputs_sha256=R.sha(R.HERE / 'analysis-inputs.json'))
    R.write(folder / 'receipt.json', receipt)
    print(name, 'exit', child.returncode, flush=True)
    assert child.returncode == 0


if __name__ == '__main__':
    execute('metrics', 'analyze_metrics.py', 'metrics-analysis.json')
    execute('profiles', 'analyze_profiles.py', 'profile-analysis.json')
