"""Root-owned matched OPC ancestor staging experiment."""
from pathlib import Path
import hashlib,json,os,shutil,subprocess,sys,time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=ROOT.parent/'litchi-target-0844'
SCRATCH=ROOT.parent/'litchi-fs-0844'
def read(p):return json.loads(Path(p).read_text())
def sha(p):
 h=hashlib.sha256()
 with Path(p).open('rb') as f:
  for chunk in iter(lambda:f.read(1048576),b''):h.update(chunk)
 return h.hexdigest()
def desc(p):
 p=Path(p);assert p.is_file() and not p.is_symlink()
 return dict(path=str(p),bytes=p.stat().st_size,sha256=sha(p))
def write(p,value):
 p=Path(p);p.parent.mkdir(parents=True,exist_ok=True)
 with p.open('x') as f:json.dump(value,f,indent=2,sort_keys=True);f.write('\n')
def output(args):return subprocess.check_output(args,cwd=ROOT,text=True).strip()
def env():
 e=os.environ.copy();assert not e.get('RUSTFLAGS') and not e.get('CARGO_ENCODED_RUSTFLAGS')
 e.update(CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',CARGO_PROFILE_RELEASE_LTO='true',CARGO_PROFILE_RELEASE_OPT_LEVEL='3',CARGO_PROFILE_RELEASE_CODEGEN_UNITS='1',CARGO_PROFILE_RELEASE_DEBUG='1',RUSTDOCFLAGS='-D warnings',PYTHONDONTWRITEBYTECODE='1',LC_ALL='C',TZ='UTC')
 return e
def source():
 names=output(['git','ls-files','--cached','--others','--exclude-standard','crates','tools','Cargo.toml','Cargo.lock','.cargo','clippy.toml','rustfmt.toml','rust-toolchain.toml']).splitlines()
 return {n:sha(ROOT/n) for n in names}
def check():
 origin=read(P/'origin.json');assert output(['git','rev-parse','HEAD'])==origin['base']
 for group in ['normative','unrelated']:
  assert all(sha(ROOT/n)==h for n,h in origin[group].items())
 for root in [TARGET,SCRATCH]:assert read(root/'.owner-0844.json')==dict(packet=str(P),path=str(root))
def run(label,argv):
 check();dest=P/'commands'/label;dest.mkdir(parents=True,exist_ok=False)
 start=time.time();write(dest/'started.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,environment={k:env().get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_PROFILE_RELEASE_LTO','CARGO_PROFILE_RELEASE_OPT_LEVEL','CARGO_PROFILE_RELEASE_CODEGEN_UNITS','CARGO_PROFILE_RELEASE_DEBUG','RUSTFLAGS']}))
 with (dest/'output.log').open('xb') as f:code=subprocess.run(argv,cwd=ROOT,env=env(),stdout=f,stderr=subprocess.STDOUT).returncode
 write(dest/'receipt.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,finished_unix=time.time(),exit_code=code,log=desc(dest/'output.log')))
 print(label,code,flush=True)
 if code:print((dest/'output.log').read_text()[-4000:],flush=True)
 return code

if __name__=='__main__':
 assert sys.argv[1]=='run'
 sys.exit(run(sys.argv[2],sys.argv[3:]))
