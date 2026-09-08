#!/usr/bin/env python3
"""Run one frozen fresh-process streaming measurement; serialize with cpu.lock."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def main():
    lane = sys.argv[1]
    protocol = json.loads((ROOT / 'protocol.json').read_text())
    assert sha(Path(__file__)) == protocol['capture_driver_sha256']
    selected = next(item for item in protocol['order'] if item['lane'] == lane)
    build = json.loads((ROOT / 'build.json').read_text())
    mode, shape = selected['mode'], selected['shape']
    binary, tree = Path(build['binaries'][mode]['path']), Path(build['build_path'])
    def check():
        assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=tree, text=True).strip() == build['revision']
        assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=tree).strip()
        assert sha(binary) == build['binaries'][mode]['sha256']
        assert sha(ROOT / build['source_manifest']['path']) == build['source_manifest']['sha256']
        for path, digest in json.loads((ROOT / build['source_manifest']['path']).read_text()).items():
            assert sha(tree / path) == digest, path
        for path, digest in build['fixtures'].items():
            assert sha(tree / path) == digest, path
    check()
    output = ROOT / 'captures' / lane
    output.mkdir(parents=True, exist_ok=False)
    argv = ['taskset', '-c', str(protocol['cpu']), '/usr/bin/time', '-v', '-o', str(output / 'resource.log'), str(binary),
            '--workers', str(protocol['workers']), '--warmup', str(protocol['warmups']), '--samples', str(protocol['samples']),
            '--case', protocol['selector'], '--semantic-shape', shape, '--json', str(output / 'report.json'),
            '--corpus-manifest', str(output / 'corpus-catalog.json')]
    env = dict(os.environ, RUSTUP_TOOLCHAIN='1.98.1', CARGO_PROFILE_RELEASE_DEBUG='1',
               RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes', DEBUGINFOD_URLS='', LC_ALL='C', PYTHONDONTWRITEBYTECODE='1')
    record = dict(schema='litchi-0474-capture-v1', **selected, revision=build['revision'], binary_sha256=sha(binary),
                  build_sha256=sha(ROOT / 'build.json'), driver_sha256=sha(Path(__file__)), protocol_sha256=sha(ROOT / 'protocol.json'),
                  argv=argv, cwd=str(tree), samples=protocol['samples'], warmups=protocol['warmups'],
                  environment={k: env[k] for k in ['RUSTUP_TOOLCHAIN', 'RUSTFLAGS', 'CARGO_PROFILE_RELEASE_DEBUG', 'DEBUGINFOD_URLS', 'LC_ALL']},
                  started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(), clean_before=True)
    (output / 'started.json').write_text(json.dumps(record, indent=2, sort_keys=True)+'\n')
    with (output / 'stdout.log').open('xb') as out, (output / 'stderr.log').open('xb') as err:
        result = subprocess.run(argv, cwd=tree, env=env, stdout=out, stderr=err)
    check()
    record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(), clean_after=True, binary_unchanged=True,
                  artifacts={p.name: dict(sha256=sha(p), bytes=p.stat().st_size) for p in output.iterdir() if p.is_file()})
    (output / 'receipt.json').write_text(json.dumps(record, indent=2, sort_keys=True)+'\n')
    print(lane, result.returncode, flush=True)
    raise SystemExit(result.returncode)

if __name__ == '__main__':
    main()
