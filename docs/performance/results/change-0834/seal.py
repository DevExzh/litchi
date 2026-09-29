"""Seal this batch's exact owned paths and verify their committed Git blobs."""
import hashlib,subprocess,sys
import driver_v3 as d

P=d.P
BASE=d.read(P/'origin.json')['base']
OUTSIDE=['docs/performance/0834-aligned-source-harness-repair.md',
 *['docs/performance/'+n+'.md' for n in ['BASELINE','CRUD_COVERAGE','GOAL_AUDIT','HOTSPOTS','REPORT']],
 *sorted(d.ALLOWED)]
def paths():
 names=set(OUTSIDE)
 for path in P.rglob('*'):
  assert not path.is_symlink()
  if path.is_file() and path.name!='seal.json':
   assert '__pycache__' not in path.parts and path.suffix!='.pyc'
   names.add(str(path.relative_to(d.ROOT)))
 return sorted(names)
def verify(committed=False):
 seal=d.read(P/'seal.json');assert seal['base']==BASE
 assert seal['files']=={n:d.sha(d.ROOT/n) for n in paths()}
 origin=d.read(P/'origin.json')
 assert all(d.sha(d.ROOT/n)==h for group in ['normative','unrelated'] for n,h in origin[group].items())
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 if committed:
  head=d.output(['git','rev-parse','HEAD']);assert d.output(['git','rev-parse','HEAD^'])==BASE
  changed=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r',head]).splitlines())
  expected=set(seal['files'])|{str((P/'seal.json').relative_to(d.ROOT))}
  assert changed==expected
  for name in expected:
   blob=subprocess.check_output(['git','show',f'{head}:{name}'],cwd=d.ROOT)
   assert hashlib.sha256(blob).hexdigest()==d.sha(d.ROOT/name)
 print('seal PASS:',len(seal['files'])+1,'owned paths')
if __name__=='__main__':
 if sys.argv[1:]==['create']:
  assert d.output(['git','rev-parse','HEAD'])==BASE
  assert d.read(P/'audit.json')['status']=='pass' and d.read(P/'cleanup.json')['status']=='pass'
  assert not (P/'capture.json').exists() and not (P/'capture-started.json').exists()
  d.write(P/'seal.json',dict(base=BASE,files={n:d.sha(d.ROOT/n) for n in paths()},disposition='harness_repair_qualified',formal_reports=0,performance_claim='none'))
  verify()
 else:
  assert sys.argv[1:] in (['verify'],['verify','--committed'])
  verify('--committed' in sys.argv)
