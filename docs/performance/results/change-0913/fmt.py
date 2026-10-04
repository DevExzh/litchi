"""Require formatting parity with the frozen baseline and clean new tests."""
import difflib
import json
from pathlib import Path
import subprocess
import tempfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
SOURCE = Path('/tmp/litchi-0913-src')
PARENT = '66b944b07fcf4c6fa37f10e7ed83a02417455f57'

def differences(data):
    proc = subprocess.run(['rustfmt','--emit','stdout','--edition','2021','--config-path',str(SOURCE/'rustfmt.toml')],input=data,capture_output=True,check=True)
    return list(difflib.unified_diff(data.decode().splitlines(),proc.stdout.decode().splitlines(),fromfile='original',tofile='formatted'))

def main():
    manifest=json.loads((P/'source-v1.json').read_text())
    result=[]
    names=manifest['overlay_paths']+['crates/soapberry-zip/tests/owned_direct_adversarial.rs']
    for name in names:
        if not name.endswith('.rs'): continue
        current=(SOURCE/name).read_bytes()
        old=subprocess.run(['git','show',PARENT+':'+name],cwd=ROOT,capture_output=True)
        before=differences(old.stdout) if old.returncode==0 else []
        after=differences(current)
        assert after==before, (name,after,before)
        result.append({'path':name,'unchanged_baseline_format_debt':bool(before),'diff':after})
    print(json.dumps({'status':'pass','scope':'No added formatting violations; baseline formatting debt explicitly retained','files':result},indent=2))

if __name__=='__main__': main()
