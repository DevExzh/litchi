#!/usr/bin/env python3
"""Build and retain an isolated measurement executable under the gate lock."""
import os
import shutil
import subprocess
import sys

from common import ROOT, REPO, TEMP, ENV, ENV_KEYS, meta, now, write
from gate import snapshot


def main():
    instrumentation, = sys.argv[1:]
    assert instrumentation in ('normal', 'allocator')
    source = snapshot()
    started = now()
    argv = ['cargo', 'build', '--release', '--locked', '--manifest-path',
            'tools/perf-baseline/Cargo.toml', '--bin', 'xml_stream_audit']
    if instrumentation == 'allocator':
        argv.extend(['--features', 'allocator-metrics'])
    subprocess.run(argv, cwd=REPO, env=ENV, check=True)
    assert snapshot() == source, 'source changed during measurement build'
    origin = REPO / 'tools/perf-baseline/target/release/xml_stream_audit'
    destination = TEMP / f'{instrumentation}-accepted' / 'xml_stream_audit'
    destination.parent.mkdir(parents=True, exist_ok=True)
    with origin.open('rb') as src, destination.open('xb') as dst:
        shutil.copyfileobj(src, dst)
    os.chmod(destination, 0o755)
    assert meta(origin) == meta(destination)
    (ROOT / 'builds').mkdir(exist_ok=True)
    write(ROOT / 'builds' / f'{instrumentation}-accepted.json', {
        'schema': 'xml-stream-audit-build-v1',
        'instrumentation': instrumentation,
        'source_snapshot': source,
        'argv': argv,
        'git_revision': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip(),
        'environment': {key: ENV[key] for key in ENV_KEYS},
        'started_utc': started,
        'finished_utc': now(),
        'binary': {'path': str(destination), **meta(destination)},
    })


if __name__ == '__main__':
    main()
