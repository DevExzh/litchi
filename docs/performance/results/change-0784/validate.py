"""Sealed offline source, raw-profile and native control replay."""
import hashlib,subprocess
from pathlib import Path
import custody as c
import native_analysis as n

def main():
 files={str(p.relative_to(c.P)):c.sha(p) for p in c.P.rglob('*') if p.is_file() and p.name!='seal.json'}
 assert files==c.read(c.P/'seal.json')['files'],'seal differs'
 for name,digest in c.read(c.P/'build/frozen-inputs.json').items():assert c.sha(c.P/name)==digest,name
 origin=c.read(c.P/'origin.json');source=c.read(c.P/'build/source.json');assert source['revision']==origin['base']
 names=subprocess.check_output(['git','ls-tree','-r','--name-only',origin['base'],'--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],text=True,cwd=c.ROOT).splitlines()
 assert set(names)==set(source['files'])
 expected=source['files']|c.read(c.P/'architecture-inputs.json')
 request=''.join(f"{origin['base']}:{name}\n" for name in expected).encode()
 data=subprocess.run(['git','cat-file','--batch'],input=request,stdout=subprocess.PIPE,check=True,cwd=c.ROOT).stdout;offset=0
 for name,digest in expected.items():
  end=data.index(b'\n',offset);header=data[offset:end].split();assert len(header)==3 and header[1]==b'blob'
  size=int(header[2]);offset=end+1;assert hashlib.sha256(data[offset:offset+size]).hexdigest()==digest,name
  offset+=size;assert data[offset:offset+1]==b'\n';offset+=1
 assert offset==len(data)
 inheritance=c.read(c.P/'inheritance.json');old=c.P.parent/inheritance['packet']
 for name,digest in inheritance['files'].items():
  assert c.sha(old/name)==digest==c.read(old/'seal.json')['files'][name]
  original=subprocess.check_output(['git','show',f"{inheritance['historical_commit']}:docs/performance/results/{inheritance['packet']}/{name}"],cwd=c.ROOT)
  assert hashlib.sha256(original).hexdigest()==digest
 for packet,commit in [('change-0780','59bdb64f16'),('change-0783','3346ed5139')]:
  original=subprocess.check_output(['git','show',f'{commit}:docs/performance/results/{packet}/seal.json'],cwd=c.ROOT)
  assert hashlib.sha256(original).hexdigest()==c.sha(c.P.parent/packet/'seal.json')
 build=c.read(c.P/'build/build.json');assert len(build['commands'])==2
 assert c.read(c.P/'build/commands.json')==build['commands']
 assert build['environment']['CARGO_BUILD_JOBS']=='2' and build['environment']['CARGO_INCREMENTAL']=='0'
 assert set(build['probe'])=={str(f.relative_to(c.P)) for f in (c.P/'probe-src').rglob('*') if f.is_file()}
 for name,digest in build['probe'].items():assert c.sha(c.P/name)==digest,name
 for row,extra in zip(build['commands'],[[],['--features','capture-profile']]):
  assert row['exit_code']==0
  expected=['cargo','build','--offline','--locked','--release','--manifest-path',str(c.P/'probe-src/Cargo.toml'),*extra]
  actual=[str(n.relocate(x)) if '/change-0784/' in x else x for x in row['command']]
  assert actual==expected
  n.artifact(row['log'])
 q=c.read(c.P/'quality.json')['format'];assert q['exit_code']==0;n.artifact(q['log'])
 for script,args in [('native_analysis.py',['--check']),('profile_analysis.py',['--check']),('perf_analysis.py',['--check'])]:
  subprocess.run(['python3','-B',str(c.P/script),*args],check=True)
 print('0784 sealed source and all observer/native analyses replay PASS')
if __name__=='__main__':main()
