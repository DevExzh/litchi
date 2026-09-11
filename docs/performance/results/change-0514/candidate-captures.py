"""Capture the frozen candidate pilots and allocation/profile admission evidence."""
import subprocess
import sys
from run import HERE, REPO

def run(script, *args):
    subprocess.run([sys.executable, '-B', str(HERE / script), *args], cwd=REPO, check=True)

for lane in ('preflight', 'pilot'):
    run('capture.py', 'after', lane)
for lane in ('preflight', 'pilot'):
    run('guard-run.py', 'after', lane)
for lane in ('allocator-r1', 'allocator-r2'):
    run('capture.py', 'after', lane)
for lane in ('allocator-r1', 'allocator-r2'):
    run('guard-run.py', 'after', lane)
run('capture.py', 'after', 'profile')
