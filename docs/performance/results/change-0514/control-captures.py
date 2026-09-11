"""Run frozen control observations serially before applying the candidate."""
import subprocess,sys
from run import HERE
for lane in ['preflight','pilot','r1','allocator-r1','allocator-r2','profile']:
    subprocess.run([sys.executable,'-B',str(HERE/'capture.py'),'before',lane],check=True)
