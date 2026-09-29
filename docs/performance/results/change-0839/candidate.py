"""Root-only candidate application/restoration, retaining exact source files."""
import driver as d
import sys
mode=sys.argv[1];d.check();names=['crates/litchi-opc/src/source_backed.rs','crates/litchi-opc/src/source_backed/batch.rs','crates/litchi-opc/tests/source_backed_batch.rs']
if mode=='apply':
 assert d.source()==d.read(d.P/'freeze-before.json')['source']
 for n in names:assert (d.ROOT/n).read_bytes()==(d.P/'source-before'/n).read_bytes()
 assert d.run('candidate-apply-check',['git','apply','--check',str(d.P/'candidate.patch')])==0
 assert d.run('candidate-apply',['git','apply',str(d.P/'candidate.patch')])==0
 inventory=d.source();before=d.read(d.P/'freeze-before.json')['source']
 assert {n for n in inventory if inventory[n]!=before[n]}==set(names)
 for n in names:
  target=d.P/'source-after'/n;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((d.ROOT/n).read_bytes())
 d.write(d.P/'candidate-source.json',dict(source=inventory,changed=names,patch=d.desc(d.P/'candidate.patch')))
elif mode=='restore':
 assert d.source()==d.read(d.P/'candidate-source.json')['source']
 for n in names:(d.ROOT/n).write_bytes((d.P/'source-before'/n).read_bytes())
 assert d.source()==d.read(d.P/'freeze-before.json')['source']
 d.write(d.P/'restored-source.json',dict(source=d.source(),reason='fresh frozen gates did not admit candidate'))
else:raise AssertionError(mode)
print(mode+' PASS')
