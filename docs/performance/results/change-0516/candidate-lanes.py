"""Run explicitly selected candidate phases serially, stopping on any failure."""
import argparse
import subprocess
import sys

from run import HERE, REPO

parser = argparse.ArgumentParser()
parser.add_argument('phase', choices=['build', 'preflight', 'pilot', 'allocator', 'profile'])
args = parser.parse_args()
if args.phase == 'build':
    commands = [['run.py', 'after', kind] for kind in ['normal', 'allocator']]
    for probe in ['guard', 'fallback']:
        commands.extend([['guard-run.py', '--probe', probe, 'after', lane]
                         for lane in ['build', 'build-allocator']])
elif args.phase == 'profile':
    commands = [['capture.py', 'after', 'profile'], ['annotate.py', 'after']]
else:
    lanes = ['allocator-r1', 'allocator-r2'] if args.phase == 'allocator' else [args.phase]
    commands = []
    for lane in lanes:
        commands.append(['capture.py', 'after', lane])
        commands.extend([['guard-run.py', '--probe', probe, 'after', lane]
                         for probe in ['guard', 'fallback']])
for script, *arguments in commands:
    subprocess.run([sys.executable, '-B', str(HERE / script), *arguments],
                   cwd=REPO, check=True)
