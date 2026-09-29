"""Apply the archived rejected representation for diagnosis, or restore it."""
import driver as d
import shutil,sys,subprocess
mode=sys.argv[1];d.check();baseline=d.read(d.P/'source-before.json')['source'];names=d.read(d.P/'candidate-origin.json')['changed']
if mode=='apply':
 assert d.source()==baseline
 subprocess.run(['git','apply','--check',str(d.P/'candidate.patch')],cwd=d.ROOT,check=True)
 subprocess.run(['git','apply',str(d.P/'candidate.patch')],cwd=d.ROOT,check=True)
 source=d.source();assert {n for n in set(baseline)|set(source) if baseline.get(n)!=source.get(n)}==set(names)
 for n in names:
  dest=d.P/'source-after'/n;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(d.ROOT/n,dest)
 d.write(d.P/'source-after.json',dict(source=source,changed=names))
elif mode=='restore':
 assert d.source()==d.read(d.P/'source-after.json')['source']
 for n in names:shutil.copy2(d.P/'source-before'/n,d.ROOT/n)
 assert d.source()==baseline
 d.write(d.P/'restored.json',dict(status='pass',source=baseline))
else:raise AssertionError(mode)
print(mode,'PASS')
