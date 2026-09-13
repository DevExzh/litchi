"""Run deterministic analyses serially, preserving every attempt and output."""
from pathlib import Path
import datetime
import hashlib
import json
import subprocess
import sys
import time

B = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def execute(name, script, args, output, stdout_report=False):
    folder = B / 'analysis-runs' / name
    folder.mkdir(parents=True, exist_ok=False)
    command = [sys.executable, '-B', str(B / script), *args]
    destination = B / output
    assert not destination.exists(), destination
    if not stdout_report:
        command += ['--output', str(destination)]
    inputs = {str(p.relative_to(B)): sha(p) for p in B.rglob('*')
              if p.is_file() and 'analysis-runs' not in p.parts
              and '__pycache__' not in p.parts}
    (folder / 'inputs.json').write_text(json.dumps(inputs, indent=2, sort_keys=True) + '\n')
    start, tick = now(), time.monotonic()
    with (folder / 'stdout').open('x') as out, (folder / 'stderr').open('x') as err:
        result = subprocess.run(command, cwd=B.parents[3], stdout=out, stderr=err)
    if stdout_report and result.returncode == 0:
        destination.write_bytes((folder / 'stdout').read_bytes())
    receipt = {'command': command, 'start_utc': start, 'end_utc': now(),
               'seconds': time.monotonic() - tick, 'exit_code': result.returncode,
               'script_sha256': sha(B / script), 'plan_sha256': sha(B / 'plan.json'),
               'inputs_sha256': sha(folder / 'inputs.json'),
               'artifacts': {p.name: sha(p) for p in folder.iterdir() if p.is_file()},
               'output': output, 'output_sha256': sha(destination) if destination.exists() else None}
    (folder / 'receipt.json').write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n')
    print(name, 'exit', result.returncode, flush=True)
    assert result.returncode == 0, name


if __name__ == '__main__':
    for script, report in [('analyze.py', 'analysis.json'),
                           ('analyze_profiles.py', 'profile-analysis.json'),
                           ('analyze_hardware.py', 'hardware-analysis.json'),
                           ('instruction_analysis.py', 'instruction-analysis.json')]:
        for stage in ['baseline', 'candidate']:
            execute(stage + '-' + script.removesuffix('.py'), script,
                    ['--stage', stage], stage + '/' + report)
        if script != 'analyze_hardware.py':
            output = {'analyze.py': 'comparison.json',
                      'analyze_profiles.py': 'profile-comparison.json',
                      'instruction_analysis.py': 'instruction-analysis-comparison.json'}[script]
            execute('compare-' + script.removesuffix('.py'), script, ['--compare'], output)
    execute('guard', 'guard_analysis.py', [], 'guard-analysis.json', True)
