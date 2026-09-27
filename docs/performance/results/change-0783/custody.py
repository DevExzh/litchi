"""0783 source and artifact custody helpers, no benchmark execution."""
from pathlib import Path
import hashlib
import json
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path('/home/zhuhe/code/litchi-target-0783')

def sha(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for b in iter(lambda:f.read(1<<20),b''):h.update(b)
    return h.hexdigest()

def read(path):return json.loads(Path(path).read_text())
def write(path,value):Path(path).write_text(json.dumps(value,indent=2,sort_keys=True)+'\n')
def artifact(path):
    path=Path(path)
    return {'path':str(path),'bytes':path.stat().st_size,'sha256':sha(path)}
def source():
    names=subprocess.check_output(['git','ls-files','-z','--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=ROOT).decode().split('\0')
    return {'revision':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'files':{n:sha(ROOT/n) for n in names if n}}
