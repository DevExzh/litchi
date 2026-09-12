"""Generate the complete report matrix from finished raw captures."""
from pathlib import Path
import subprocess
import sys

HERE = Path(__file__).resolve().parent


def invoke(script, *args):
    subprocess.run([sys.executable,'-B',str(HERE/script),*args],check=True)


for stage in ['baseline','candidate']:
    invoke('analyze.py','--stage',stage)
invoke('analyze.py','--compare')
for stage in ['baseline','candidate']:
    invoke('analyze_profiles.py','--stage',stage)
invoke('analyze_profiles.py','--compare')
for stage in ['baseline','candidate']:
    invoke('analyze_hardware.py','--stage',stage)
invoke('analyze_assembly.py')
