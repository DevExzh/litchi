#!/usr/bin/env python3
"""Retain validation commands with compact content-addressed source custody."""
from pathlib import Path
import hashlib
import json
import subprocess
import sys
from common import ROOT, REPO, ENV, ENV_KEYS, sha, meta, now, write


def snapshot():
    names = subprocess.check_output(['git', 'ls-files', '-co', '--exclude-standard', '-z'], cwd=REPO).decode().split('\0')
    sources = {p: sha(REPO / p) for p in sorted(set(names))
               if p and Path(p).suffix in ('.rs', '.toml', '.lock') and (REPO / p).is_file()}
    encoded = (json.dumps(sources, indent=2, sort_keys=True) + '\n').encode()
    digest = hashlib.sha256(encoded).hexdigest()
    directory = ROOT / 'validation-sources'; directory.mkdir(exist_ok=True)
    path = directory / f'{digest}.json'
    if not path.exists():
        with path.open('xb') as stream: stream.write(encoded)
    assert sha(path) == digest
    return dict(path=path.relative_to(ROOT).as_posix(), sha256=digest, files=len(sources))


def main():
    label, *argv = sys.argv[1:]
    assert label and '/' not in label and argv
    prefix = ROOT / 'validation' / label
    record = dict(argv=argv, cwd=str(REPO), environment={k: ENV[k] for k in ENV_KEYS},
        driver_sha256=sha(Path(__file__)), common_sha256=sha(ROOT / 'common.py'),
        started_utc=now(), source_before=snapshot())
    write(prefix.with_suffix('.started.json'), record)
    with prefix.with_suffix('.stdout').open('xb') as out, prefix.with_suffix('.stderr').open('xb') as err:
        result = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err)
    record.update(exit_code=result.returncode, finished_utc=now(), source_after=snapshot(),
        artifacts={prefix.with_suffix(s).name: meta(prefix.with_suffix(s)) for s in ('.stdout', '.stderr')})
    record['source_unchanged'] = record['source_before'] == record['source_after']
    write(prefix.with_suffix('.json'), record)
    print(label, result.returncode, 'source unchanged', record['source_unchanged'], flush=True)
    raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
