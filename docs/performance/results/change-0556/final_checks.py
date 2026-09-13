"""Run final documentation and boundary checks against restored source."""
import json
import subprocess
import time

from quality import HERE, REPO, now, sha, verify_source, write


def main():
    root = HERE / 'final-checks'
    root.mkdir(exist_ok=False)
    manifest = json.loads((HERE / 'baseline-source.json').read_text())
    commands = {
        'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
        'claims': ['python3', '-B', 'tools/check_perf_claims.py', '--registry',
                   'docs/performance/claim-registry-v1.json', '--repo-root', '.',
                   '--evidence-root', '.', '--mode', 'strict'],
        'format': ['cargo', 'fmt', '--all', '--', '--check'],
    }
    write(root / 'inputs.json', dict(commands=commands,
          driver_sha256=sha(HERE / 'final_checks.py'),
          source_manifest_sha256=sha(HERE / 'baseline-source.json')))
    rows = []
    for name, command in commands.items():
        verify_source(manifest)
        start, tick = now(), time.monotonic()
        with (root / (name + '.stdout')).open('x') as out, (root / (name + '.stderr')).open('x') as err:
            result = subprocess.run(command, cwd=REPO, stdout=out, stderr=err)
        verify_source(manifest)
        receipt = dict(command=command, exit_code=result.returncode,
                       start_utc=start, end_utc=now(), seconds=time.monotonic() - tick,
                       source_stable=True, inputs_sha256=sha(root / 'inputs.json'),
                       stdout_sha256=sha(root / (name + '.stdout')),
                       stderr_sha256=sha(root / (name + '.stderr')))
        write(root / (name + '.json'), receipt)
        rows.append(dict(name=name, receipt_sha256=sha(root / (name + '.json')),
                         exit_code=result.returncode))
        print(name, 'exit', result.returncode, flush=True)
        if result.returncode:
            break
    passed = len(rows) == len(commands) and all(r['exit_code'] == 0 for r in rows)
    write(root / 'result.json', dict(status='pass' if passed else 'failed', rows=rows))
    return 0 if passed else 1


if __name__ == '__main__':
    raise SystemExit(main())
