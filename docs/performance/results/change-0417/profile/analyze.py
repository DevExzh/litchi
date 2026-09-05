#!/usr/bin/env python3
"""Retain readable perf views while the exact profiled ELF is available."""
from pathlib import Path
import datetime
import json
import os
import subprocess

root = Path(__file__).resolve().parent
env = os.environ.copy()
env['DEBUGINFOD_URLS'] = ''
commands = [
    ('header.txt', ['perf', 'report', '--header-only', '-i', str(root / 'record.data')]),
    ('self.txt', ['perf', 'report', '--stdio', '--no-inline', '--no-children',
                  '--percent-limit', '0.1', '--sort', 'symbol,dso', '-i', str(root / 'record.data')]),
    ('children.txt', ['perf', 'report', '--stdio', '--no-inline', '--children',
                      '--percent-limit', '1', '--sort', 'symbol,dso', '-i', str(root / 'record.data')]),
    ('script.txt', ['perf', 'script', '--no-inline', '-F', 'comm,pid,tid,time,period,event,ip,sym,dso',
                    '-i', str(root / 'record.data')]),
]
records = []
for name, argv in commands:
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with (root / name).open('wb') as stdout, (root / (name + '.stderr')).open('wb') as stderr:
        process = subprocess.run(argv, env=env, stdout=stdout, stderr=stderr)
    records.append(dict(argv=argv, started_utc=started, exit_code=process.returncode,
                        stdout=name, stderr=name + '.stderr'))
    (root / 'analysis-commands.json').write_text(json.dumps(records, indent=2) + '\n')
    if process.returncode:
        raise SystemExit(process.returncode)
