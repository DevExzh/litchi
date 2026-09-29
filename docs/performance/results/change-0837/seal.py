"""Seal only this batch; verify the exact committed owned path set afterward."""
import sys
import driver as d
seal=d.P/'seal.json'
if sys.argv[1:]==['create']:
 assert d.read(d.P/'closure.json')['status']=='pass'
 assert d.read(d.P/'cleanup.json')['status']=='pass'
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 assert d.output(['git','rev-parse','HEAD'])==d.read(d.P/'origin.json')['base']
 names=['crates/soapberry-zip/src/office.rs','crates/soapberry-zip/tests/streaming_interrupted_write.rs','docs/performance/0837-zip-interrupted-write-and-fresh-compression.md']
 names += ['docs/performance/'+n+'.md' for n in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
 for path in sorted(d.P.rglob('*')):
  assert not path.is_symlink()
  if path.is_file():
   assert path.suffix!='.pyc' and '__pycache__' not in path.parts
   names.append(str(path.relative_to(d.ROOT)))
 assert len(names)==len(set(names))
 d.write(seal,dict(base=d.read(d.P/'origin.json')['base'],disposition='zip-write-interruption-fix; fresh-compression-rejected',files={n:d.sha(d.ROOT/n) for n in sorted(names)}))
else:assert sys.argv[1:]==['check']
record=d.read(seal)
for name,h in record['files'].items():assert d.sha(d.ROOT/name)==h,name
head=d.output(['git','rev-parse','HEAD'])
if head!=record['base']:
 assert d.output(['git','rev-parse','HEAD^'])==record['base']
 actual=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())
 assert actual==set(record['files'])|{str(seal.relative_to(d.ROOT))}
 for name in actual:
  content=d.subprocess.check_output(['git','show',f'HEAD:{name}'],cwd=d.ROOT)
  assert d.hashlib.sha256(content).hexdigest()==d.sha(d.ROOT/name)
print('seal PASS',len(record['files'])+1,'owned paths',head)
