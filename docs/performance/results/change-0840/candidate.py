"""Root-only exact patch application and restoration of owned candidate files."""
import driver as d
import shutil,sys
mode=sys.argv[1];core='crates/litchi-cfb/src/writer/core.rs';test='crates/litchi-cfb/tests/writer_emission_order.rs'
d.check()
if mode=='test':
 assert d.source()==d.read(d.P/'source-before.json')['source']
 assert d.run('add-focused-test',['git','apply','--include='+test,str(d.P/'candidate.patch')])==0
 assert d.sha(d.ROOT/core)==d.read(d.P/'source-before.json')['archived'][core]
elif mode=='apply':
 assert d.sha(d.ROOT/core)==d.read(d.P/'source-before.json')['archived'][core]
 assert d.run('apply-candidate',['git','apply','--include='+core,str(d.P/'candidate.patch')])==0
 before=d.read(d.P/'source-before.json')['source'];after=d.source()
 changed=sorted(n for n in set(before)|set(after) if before.get(n)!=after.get(n));assert changed==sorted([core,test]),changed
 for n in changed:
  p=d.P/'source-after'/n;p.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(d.ROOT/n,p)
 d.write(d.P/'candidate-source.json',dict(source=after,changed=changed,patch=d.desc(d.P/'candidate.patch')))
elif mode=='restore':
 candidate=d.read(d.P/'candidate-source.json');assert d.source()==candidate['source']
 shutil.copy2(d.P/'source-before'/core,d.ROOT/core);(d.ROOT/test).unlink()
 assert d.source()==d.read(d.P/'source-before.json')['source']
 d.write(d.P/'restoration.json',dict(status='pass',source=d.source()))
else:raise AssertionError(mode)
print(mode,'PASS')
