"""Capture remaining control pilot, allocation repeats, and profile serially."""
import subprocess,sys
from run import HERE,REPO
commands=[['guard-run.py','--probe','fallback','before','pilot']]
for lane in ['allocator-r1','allocator-r2']:
    commands.extend([['capture.py','before',lane],['guard-run.py','before',lane],['guard-run.py','--probe','fallback','before',lane]])
commands.append(['capture.py','before','profile'])
for script,*arguments in commands:
    subprocess.run([sys.executable,'-B',str(HERE/script),*arguments],cwd=REPO,check=True)
