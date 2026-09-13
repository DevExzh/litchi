"""Preserve a strict verifier attempt without overwriting prior evidence."""
from pathlib import Path
import argparse
import datetime
import hashlib
import json
import shutil
import subprocess
import time

B = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--folder', type=Path, required=True)
    parser.add_argument('--component', default='precleanup')
    args = parser.parse_args()
    folder = args.folder.resolve()
    assert not (B / 'SHA256SUMS').exists() or not folder.is_relative_to(B)
    folder.mkdir(parents=True, exist_ok=False)
    script = B / 'verify.py'
    shutil.copy2(script, folder / 'verify.py')
    before = sha(script)
    command = ['python3', '-B', str(script), '--component', args.component, '--strict']
    start, tick = now(), time.monotonic()
    with (folder / 'stdout').open('x') as out, (folder / 'stderr').open('x') as err:
        result = subprocess.run(command, cwd=B.parents[3], stdout=out, stderr=err)
    receipt = {'command': command, 'start_utc': start, 'end_utc': now(),
               'seconds': time.monotonic() - tick, 'exit_code': result.returncode,
               'script_sha256_before': before, 'script_sha256_after': sha(script),
               'artifacts': {p.name: sha(p) for p in folder.iterdir() if p.is_file()}}
    (folder / 'receipt.json').write_text(json.dumps(receipt, indent=2) + '\n')
    print('Strict', args.component, 'exit', result.returncode, 'verifier stable', before == sha(script))
    report = json.loads((folder / 'stdout').read_text())
    print(json.dumps({key: report[key] for key in ['status', 'error'] if key in report}))
    assert before == sha(script), 'verifier changed during its attempt'
    raise SystemExit(result.returncode)
