"""Separate whole-process PPTX CPU diagnostics, after formal captures."""
from pathlib import Path
import datetime
import hashlib
import json
import os
import subprocess
import sys

repo = Path('/home/zhuhe/code/litchi')
root = repo / 'docs/performance/results/change-0417'
tree = Path('/tmp/litchi-goal-0417-worktree')
out = root / 'profile'
out.mkdir(exist_ok=True)
mode = sys.argv[1]
assert mode in ('record', 'stat')
build = json.loads((root / 'build-identity.json').read_text())
capture = json.loads((root / 'capture.json').read_text())
assert len(capture['runs']) == 120 and all(r['exit_code'] == 0 for r in capture['runs'])
assert subprocess.check_output(['git', 'status', '--porcelain'], cwd=tree, text=True) == ''
binary = build['binaries']['normal']
assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest() == binary['sha256']
report = out / (mode + '.json')
catalog = out / (mode + '.catalog.json')
args = [binary['path'], '--case', 'pptx_cross_copy_media_rich',
        *json.loads((root / 'matrix.json').read_text())['common_flags'],
        '--samples', '20', '--warmup', '3', '--json', str(report),
        '--corpus-manifest', str(catalog)]
if mode == 'record':
    prefix = ['perf', 'record', '--no-buildid-cache', '-e', 'cycles:u', '-F', '999',
              '--call-graph', 'fp,127', '-o', str(out / 'record.data'), '--']
else:
    prefix = ['perf', 'stat', '--no-big-num', '-x,', '-e',
              'cycles,instructions,branches,branch-misses,L1-dcache-loads,L1-dcache-load-misses,LLC-loads,LLC-load-misses,page-faults,task-clock',
              '-o', str(out / 'stat.csv'), '--']
argv = ['taskset', '-c', '2', *prefix, *args]
journal = out / (mode + '-command.json')
assert not journal.exists() and not report.exists(), 'refuse overwrite'
env = os.environ.copy()
env['DEBUGINFOD_URLS'] = ''
started = datetime.datetime.now(datetime.timezone.utc).isoformat()
with (out / (mode + '.stdout')).open('wb') as stdout, (out / (mode + '.stderr')).open('wb') as stderr:
    process = subprocess.run(argv, cwd=tree, env=env, stdout=stdout, stderr=stderr)
journal.write_text(json.dumps(dict(
    revision=build['revision'], source_status='', binary_sha256=binary['sha256'],
    argv=argv, cwd=str(tree), environment_overrides={'DEBUGINFOD_URLS': ''},
    started_utc=started, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
    exit_code=process.returncode,
    scope='Whole process including corpus generation, preflight, setup, warmups, timed phases, verification and reporting; diagnostic only; no elapsed comparison',
), indent=2) + '\n')
raise SystemExit(process.returncode)
