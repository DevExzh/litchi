"""Serial current-source XML attribute integration gates; failed attempts are retained."""
from pathlib import Path
import hashlib,json,os,subprocess,time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
TARGET=Path('/home/zhuhe/code/litchi-target-0777-quality')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def census():
 names=subprocess.check_output(['git','ls-files','-z'],cwd=ROOT).decode().split('\0')
 return {n:sha(ROOT/n) for n in names if n and ((n.startswith('crates/') or n.startswith('tools/perf-baseline/')) and (n.endswith('.rs') or Path(n).name in ['Cargo.toml','Cargo.lock']) or n in ['Cargo.toml','.cargo/config.toml','rust-toolchain.toml','clippy.toml'])}|{'Cargo.lock':sha(ROOT/'Cargo.lock')}
if __name__=='__main__':
 before=census();i=0
 while (P/f'quality-{i}').exists():i+=1
 folder=P/f'quality-{i}';folder.mkdir();(folder/'source.json').write_text(json.dumps(before,indent=2)+'\n')
 packages=['litchi-crypto','litchi-docx','litchi-drawingml','litchi-formula','litchi-imgconv','litchi-ole-common','litchi-ooxml-common','litchi-opc','litchi-ppt','litchi-pptx','litchi-sign','litchi-spreadsheet-drawing','litchi-xldm','litchi-xlsb','litchi-xlsx','xml-minifier']
 owners=[arg for name in packages for arg in ['-p',name]]
 common=['--offline','--locked','--all-features']
 commands=[['cargo','fmt',*owners,'--','--check'],
 ['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--','--check'],
 ['cargo','check',*owners,*common,'--all-targets'],
 ['cargo','test',*owners,*common,'--','--test-threads=2'],
 ['cargo','clippy',*owners,*common,'--lib','--','-D','warnings'],
 ['cargo','doc',*owners,*common,'--no-deps'],
 ['cargo','check','-p','litchi','--offline','--locked','--features','doc,docx,ppt,pptx,xls,xlsb,xlsx','--all-targets'],
 ['cargo','check','--manifest-path','tools/perf-baseline/Cargo.toml','--offline','--locked','--all-targets'],
 ['python3','-B','tools/check_crate_boundaries.py']]

 env=os.environ|{'CARGO_TARGET_DIR':str(TARGET),'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','CARGO_PROFILE_DEV_DEBUG':'0','RUSTDOCFLAGS':'-D warnings','PYTHONDONTWRITEBYTECODE':'1'}
 rows=[]
 for index,cmd in enumerate(commands):
  log=folder/f'{index:02d}.log';start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=env,stdout=f,stderr=subprocess.STDOUT)
  rows.append({'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':str(log.relative_to(P)),'sha256':sha(log)})
  receipt={'rows':rows,'source_file':str((folder/'source.json').relative_to(P)),'source_sha256':sha(folder/'source.json'),'target':str(TARGET),'environment':{n:env.get(n) for n in ['CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_DEV_DEBUG','RUSTDOCFLAGS','RUSTFLAGS','RUSTUP_TOOLCHAIN']}}
  (folder/'manifest.json').write_text(json.dumps(receipt,indent=2)+'\n');print(f'gate {index+1}/{len(commands)} exit {r.returncode}',flush=True)
  assert r.returncode==0,rows[-1]
  assert census()==before,'source/config/lock changed'
 (P/'quality.json').write_text(json.dumps(receipt,indent=2)+'\n')
