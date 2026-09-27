"""Local artifact helpers for the standalone RSS accounting calibration."""
from pathlib import Path
import hashlib,json
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=Path('/home/zhuhe/code/litchi-target-0790')
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def read(path):return json.loads(Path(path).read_text())
def write(path,data):Path(path).write_text(json.dumps(data,indent=2,sort_keys=True)+'\n')
def artifact(path):
 path=Path(path)
 try:name=str(path.relative_to(P))
 except ValueError:name=str(path)
 return {'path':name,'bytes':path.stat().st_size,'sha256':sha(path)}
def resolve(a):return P/a['path']
def verify(a):
 p=resolve(a);assert p.stat().st_size==a['bytes'] and sha(p)==a['sha256'];return p
