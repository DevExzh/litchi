"""Serial source-bound OLE2 quality checks; preserve every actual attempt."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path('/home/zhuhe/litchi-goal-0555-target')


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2)
        stream.write('\n')


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def verify_source(manifest):
    for name, digest in manifest.items():
        assert sha(REPO / name) == digest, name
    assert sha(REPO / 'Cargo.lock') == json.loads(
        (HERE / 'workspace-lock.json').read_text())['sha256']


def main():
    stage, label, mode = sys.argv[1:]
    assert stage in ('baseline', 'candidate', 'final')
    assert mode in ('targeted', 'commands')
    assert label and all(c.isalnum() or c in '-_' for c in label)
    root = HERE / 'quality-attempts' / label
    root.mkdir(parents=True, exist_ok=False)
    manifest_path = HERE / stage / 'source-manifest.json'
    manifest = json.loads(manifest_path.read_text())
    plan = json.loads((HERE / 'quality-plan.json').read_text())
    environment = dict(os.environ)
    environment.update(TMPDIR=str(TARGET / 'tmp'), CARGO_TARGET_DIR=str(TARGET),
                       CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0',
                       RUSTDOCFLAGS='-D warnings')
    inputs = dict(schema='ole2_0555_quality_inputs_v1', created_utc=now(),
                  stage=stage, mode=mode, source_manifest_sha256=sha(manifest_path),
                  plan_sha256=sha(HERE / 'quality-plan.json'),
                  script_sha256=sha(Path(__file__)), commands=plan[mode],
                  workspace_lock_sha256=sha(REPO / 'Cargo.lock'))
    write(root / 'inputs.json', inputs)
    rows = []
    for name, command in plan[mode].items():
        verify_source(manifest)
        folder = root / name
        folder.mkdir()
        start, tick = now(), time.monotonic()
        with (folder / 'stdout').open('x') as out, (folder / 'stderr').open('x') as err:
            result = subprocess.run(command, cwd=REPO, env=environment,
                                    stdout=out, stderr=err)
        verify_source(manifest)
        receipt = dict(schema='ole2_0555_quality_receipt_v1', command=command,
                       start_utc=start, end_utc=now(), seconds=time.monotonic()-tick,
                       exit_code=result.returncode, source_stable=True,
                       source_manifest_sha256=sha(manifest_path),
                       inputs_sha256=sha(root / 'inputs.json'),
                       environment={k: environment.get(k) for k in (
                           'TMPDIR', 'CARGO_TARGET_DIR', 'CARGO_BUILD_JOBS',
                           'CARGO_INCREMENTAL', 'RUSTDOCFLAGS', 'RUSTFLAGS',
                           'CARGO_ENCODED_RUSTFLAGS', 'LD_PRELOAD')},
                       artifacts={p: sha(folder / p) for p in ('stdout', 'stderr')})
        write(folder / 'receipt.json', receipt)
        rows.append(dict(name=name, path=str(folder.relative_to(HERE)),
                         receipt_sha256=sha(folder / 'receipt.json'),
                         exit_code=result.returncode))
        print(label, name, 'exit', result.returncode, flush=True)
        if result.returncode:
            break
    passed = len(rows) == len(plan[mode]) and all(row['exit_code'] == 0 for row in rows)
    write(root / 'result.json', dict(schema='ole2_0555_quality_v1',
          status='pass' if passed else 'failed', stage=stage, mode=mode,
          source_manifest_sha256=sha(manifest_path),
          inputs_sha256=sha(root / 'inputs.json'), completed_utc=now(), rows=rows))
    return 0 if passed else 1


if __name__ == '__main__':
    raise SystemExit(main())
