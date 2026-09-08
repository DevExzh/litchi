#!/usr/bin/env python3
"""Run one serialized check and retain exact command, logs and source custody."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def main():
    name, *argv = sys.argv[1:]
    output = ROOT / 'validation'
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_BUILD_JOBS='4',
               CARGO_INCREMENTAL='0', PYTHONDONTWRITEBYTECODE='1')
    keys = ['RUSTUP_TOOLCHAIN', 'CARGO_BUILD_JOBS', 'CARGO_INCREMENTAL',
            'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG', 'RUSTDOCFLAGS']
    sources = {p: sha(REPO / p) for p in subprocess.check_output(
        ['git', 'ls-files'], cwd=REPO, text=True).splitlines()
        if Path(p).suffix in ('.rs', '.toml', '.lock') and (REPO / p).is_file()}
    for p in (REPO / 'tools/perf-baseline/src').glob('docx_streaming_create.rs'):
        sources[p.relative_to(REPO).as_posix()] = sha(p)
    record = dict(argv=argv, cwd=str(REPO), source_hashes=sources,
                  environment={k: env[k] for k in keys if k in env},
                  started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
    with (output / (name + '.stdout')).open('xb') as out, (output / (name + '.stderr')).open('xb') as err:
        result = subprocess.run(argv, cwd=REPO, env=env, stdout=out, stderr=err)
    record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                  artifacts={name + suffix: dict(sha256=sha(output / (name + suffix)), bytes=(output / (name + suffix)).stat().st_size)
                             for suffix in ('.stdout', '.stderr')})
    with (output / (name + '.json')).open('x') as stream:
        json.dump(record, stream, indent=2, sort_keys=True); stream.write('\n')
    print(name, result.returncode, flush=True)
    raise SystemExit(result.returncode)

if __name__ == '__main__':
    main()
