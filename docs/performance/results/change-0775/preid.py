"""Separate semantic reference capture; run only after paired native lane ends."""
from pathlib import Path
import hashlib,json,os,subprocess,time,shutil
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/home/zhuhe/code/litchi-worktrees/0771-preid-src')
TARGET=Path('/home/zhuhe/code/litchi-target-0775-preid')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def census():
 names=subprocess.check_output(['git','ls-files','-z','crates','Cargo.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=SOURCE).decode().split('\0')
 return {n:sha(SOURCE/n) for n in names if n}
if __name__=='__main__':
 out=P/'preid-0';assert not out.exists();out.mkdir();before=census();write(out/'source.json',before)
 ref=subprocess.check_output(['git','rev-parse','HEAD'],cwd=SOURCE).decode().strip();assert ref.startswith('bdea6e8b86')
 subprocess.run(['git','diff','--quiet','HEAD','--','crates','Cargo.toml'],cwd=SOURCE,check=True)
 (out/'src').mkdir();shutil.copy2(P/'probe-src/main.rs',out/'src/main.rs');(out/'Cargo.toml').write_text((P/'probe-src/Cargo.toml.template').read_text().replace('@SRC@',str(SOURCE)));shutil.copy2(P/'measure-2/before/Cargo.lock',out/'Cargo.lock')
 env=os.environ|{'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','RUSTUP_TOOLCHAIN':'1.95.0','CARGO_TARGET_DIR':str(TARGET)}
 rows=[]
 commands=[['cargo','build','--release','--offline','--locked','--manifest-path',str(out/'Cargo.toml')],['taskset','-c','12',str(TARGET/'release/mce-stream-probe'),'differential','--generated','4096','--aliasing','4096','--json',str(out/'differential.json'),'test-data/ooxml/docx','test-data/ooxml/xlsx','test-data/ooxml/pptx']]
 for i,cmd in enumerate(commands):
  log=out/f'{i}.log';start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
  rows.append({'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':log.name,'sha256':sha(log)});write(out/'runs.json',rows);assert r.returncode==0
  if i==0:write(out/'binary.json',{'path':str(TARGET/'release/mce-stream-probe'),'sha256':sha(TARGET/'release/mce-stream-probe'),'source':ref,'probe_sha256':sha(out/'src/main.rs'),'lock_sha256':sha(out/'Cargo.lock')})
 assert census()==before
 write(out/'complete.json',{'source_unchanged':True,'report_sha256':sha(out/'differential.json')})
