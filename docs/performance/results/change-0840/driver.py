"""Root-owned fresh CFB emission experiment."""
from pathlib import Path
import hashlib,json,os,shutil,subprocess,sys,time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=ROOT.parent/'litchi-target-0840'
SCRATCH=ROOT.parent/'litchi-fs-0840'
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
 e.update(CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',CARGO_PROFILE_RELEASE_LTO='true',RUSTDOCFLAGS='-D warnings',PYTHONDONTWRITEBYTECODE='1',LC_ALL='C',TZ='UTC')
 return e
def source():
 names=output(['git','ls-files','--cached','--others','--exclude-standard','crates','tools','Cargo.toml','Cargo.lock','.cargo','clippy.toml','rustfmt.toml','rust-toolchain.toml']).splitlines()
 return {n:sha(ROOT/n) for n in names}
def check():
 origin=read(P/'origin.json');assert output(['git','rev-parse','HEAD'])==origin['base']
 for group in ['normative','unrelated']:
  assert all(sha(ROOT/n)==h for n,h in origin[group].items())
 for root in [TARGET,SCRATCH]:assert read(root/'.owner-0840.json')==dict(packet=str(P),path=str(root))
def prepare():
 assert not TARGET.exists() and not SCRATCH.exists()
 prior=read(P.parent/'change-0839/origin.json')
 for group in ['normative','unrelated']:
  assert all(sha(ROOT/n)==h for n,h in prior[group].items())
 write(P/'origin.json',dict(base=output(['git','rev-parse','HEAD']),normative=prior['normative'],unrelated=prior['unrelated'],scope='Fresh CFB physical-order emission qualification; output bytes unchanged; iWork excluded.'))
 for root in [TARGET,SCRATCH]:root.mkdir();write(root/'.owner-0840.json',dict(packet=str(P),path=str(root)))
 write(P/'host.json',dict(rustc=output(['rustc','-Vv']),cargo=output(['cargo','-V']),cpu=Path('/proc/cpuinfo').read_text(),memory=Path('/proc/meminfo').read_text(),affinity=sorted(os.sched_getaffinity(0)),filesystem=output(['findmnt','-T',str(ROOT),'-n','-o','FSTYPE,OPTIONS']),environment={k:env().get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','CARGO_PROFILE_RELEASE_LTO','RUSTFLAGS','RUSTDOCFLAGS']}))
 assert 12 in os.sched_getaffinity(0)
 print('prepared isolated0840roots',flush=True)
def run(label,argv):
 check();dest=P/'commands'/label;dest.mkdir(parents=True,exist_ok=False)
 start=time.time();write(dest/'started.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,environment={k:env().get(k) for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_PROFILE_RELEASE_LTO','RUSTFLAGS']}))
 with (dest/'output.log').open('xb') as f:code=subprocess.run(argv,cwd=ROOT,env=env(),stdout=f,stderr=subprocess.STDOUT).returncode
 write(dest/'receipt.json',dict(argv=argv,cwd=str(ROOT),started_unix=start,finished_unix=time.time(),exit_code=code,log=desc(dest/'output.log')))
 print(label,code,flush=True)
 if code:print((dest/'output.log').read_text()[-4000:],flush=True)
 return code
def build(stage):
 check();inventory=source();probe={str(p.relative_to(P)):sha(p) for p in sorted((P/'probe').rglob('*')) if p.is_file()}
 write(P/f'freeze-{stage}.json',dict(source=inventory,probe=probe,driver=desc(P/'driver.py')))
 assert run('build-'+stage,['cargo','build','--offline','--locked','--release','--manifest-path',str(P/'probe/Cargo.toml')])==0
 assert source()==inventory
 binary=TARGET/'retained'/stage/'cfb-emission-probe';binary.parent.mkdir(parents=True,exist_ok=False);shutil.copy2(TARGET/'release/cfb-emission-probe',binary)
 write(P/f'build-{stage}.json',dict(binary=desc(binary),freeze=desc(P/f'freeze-{stage}.json'),receipt=desc(P/f'commands/build-{stage}/receipt.json')))
if __name__=='__main__':
 command=sys.argv[1]
 if command=='prepare':prepare()
 elif command=='build':build(sys.argv[2])
 elif command=='run':sys.exit(run(sys.argv[2],sys.argv[3:]))
 else:raise AssertionError(command)
