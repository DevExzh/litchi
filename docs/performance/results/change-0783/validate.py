"""Offline retained evidence validation, including after executable cleanup."""
import hashlib
import subprocess
import custody as c
import analyze

def main():
 seal=c.read(c.P/'seal.json')['files']
 actual={str(p.relative_to(c.P)):c.sha(p) for p in c.P.rglob('*') if p.is_file() and p.name!='seal.json'}
 assert actual==seal,'packet seal differs'
 for name,digest in c.read(c.P/'build/frozen-inputs.json').items():
  assert c.sha(c.P/name)==digest,name
 old=c.P.parent/'change-0780'
 inherited=c.read(c.P/'inheritance.json')
 old_seal=c.read(old/'seal.json')['files']
 historical_seal=subprocess.check_output(['git','show',f"{inherited['historical_commit']}:docs/performance/results/change-0780/seal.json"],cwd=c.ROOT)
 assert hashlib.sha256(historical_seal).hexdigest()==c.sha(old/'seal.json')
 for name,digest in inherited['files'].items():
  assert c.sha(old/name)==digest==old_seal[name],name
  data=subprocess.check_output(['git','show',f"{inherited['historical_commit']}:docs/performance/results/change-0780/{name}"],cwd=c.ROOT)
  assert hashlib.sha256(data).hexdigest()==digest,name
 quality=c.read(c.P/'quality.json')['format']
 assert quality['exit_code']==0
 assert c.sha(c.P/quality['log']['path'])==quality['log']['sha256']
 origin=c.read(c.P/'origin.json')
 source=c.read(c.P/'build/source.json')
 assert source['revision']==origin['base']
 names=subprocess.check_output(['git','ls-tree','-r','--name-only',origin['base'],'--','crates','Cargo.toml','clippy.toml','.cargo/config.toml','rust-toolchain.toml'],cwd=c.ROOT,text=True).splitlines()
 assert set(names)==set(source['files'])
 expected=source['files']|c.read(c.P/'architecture-inputs.json')
 requests=''.join(f"{origin['base']}:{name}\n" for name in expected).encode()
 objects=subprocess.run(['git','cat-file','--batch'],input=requests,stdout=subprocess.PIPE,check=True,cwd=c.ROOT).stdout
 offset=0
 for name,digest in expected.items():
  end=objects.index(b'\n',offset)
  header=objects[offset:end].split();assert len(header)==3 and header[1]==b'blob',name
  size=int(header[2]);offset=end+1
  assert hashlib.sha256(objects[offset:offset+size]).hexdigest()==digest,name
  offset+=size;assert objects[offset:offset+1]==b'\n';offset+=1
 assert offset==len(objects)
 subprocess.run(['python3','-B',str(c.P/'analyze.py'),'--check'],check=True)
 print('Sealed source, architecture, capture and analysis replay PASS')
if __name__=='__main__':main()
