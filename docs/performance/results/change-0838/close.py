"""Verify bindings, clean only marked roots, and seal the owned batch."""
import driver as d
import json,shutil,sys
REPORT='docs/performance/0838-docx-durable-inverse-physical-diagnostic.md'
INDEXES=[f'docs/performance/{n}.md' for n in ['BASELINE','HOTSPOTS','REPORT','CRUD_COVERAGE','GOAL_AUDIT']]
def local(p):
 s=str(p);marker='/docs/performance/'
 assert marker in s
 return d.ROOT/'docs/performance'/s.split(marker,1)[1]
def binding(x):
 p=local(x['path']);assert p.stat().st_size==x['bytes'] and d.sha(p)==x['sha256'],str(p)
def verify():
 origin=d.read(d.P/'origin.json')
 for group in ['normative','unrelated']:
  assert all(d.sha(d.ROOT/n)==h for n,h in origin[group].items())
 reuse=d.read(d.P/'quality-reuse.json');assert d.source()==reuse['source']
 binding(reuse['prior_quality']);binding(reuse['prior_source'])
 quality=d.read(d.P/'quality.json');freeze=d.read(d.P/'freeze-diagnostic.json')
 assert quality['source']==freeze['source']==reuse['source']
 probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
 assert quality['probe']==freeze['probe']==probe
 for x in quality['commands']:binding(x);assert d.read(local(x['path']))['exit_code']==0
 for x in d.read(d.P/'admission.json')['files']:binding(x)
 build=d.read(d.P/'build-diagnostic.json');binding(build['freeze']);binding(build['receipt']);binding(freeze['driver'])
 for x in d.read(d.P/'capture.json')['artifacts']:binding(x)
 commands=[]
 for p in sorted((d.P/'commands').iterdir()):
  r=d.read(p/'receipt.json');s=d.read(p/'started.json');assert r['argv']==s['argv'] and r['started_unix']==s['started_unix'] and r['finished_unix']>=r['started_unix'];binding(r['log'])
  assert r['exit_code']==(101 if p.name=='probe-clippy' else 0)
  commands.append(dict(label=p.name,exit_code=r['exit_code']))
 captures=[d.read(d.P/f'commands/capture-{i:02}/receipt.json') for i in range(3)]
 for i,r in enumerate(captures):
  assert r['argv']==['taskset','-c','12',build['binary']['path'],str(d.P/'fixtures/fresh-source.zip'),str(d.P/'runs'/f'run-{i:02}')]
  if i:assert captures[i-1]['finished_unix']<=r['started_unix']
 assert d.read(d.P/'analysis.json')['status']=='pass'
 return commands
command=sys.argv[1]
if command=='clean':
 d.check();commands=verify();removed=[]
 for root in [d.TARGET,d.SCRATCH]:
  assert d.read(root/'.owner-0838.json')==dict(packet=str(d.P),path=str(root))
  files=[p for p in root.rglob('*') if p.is_file()];removed.append(dict(path=str(root),files=len(files),logical_bytes=sum(p.stat().st_size for p in files)))
  if root==d.TARGET:assert d.desc(d.read(d.P/'build-diagnostic.json')['binary']['path'])==d.read(d.P/'build-diagnostic.json')['binary']
  shutil.rmtree(root);assert not root.exists()
 d.write(d.P/'cleanup.json',dict(status='pass',removed=removed))
 d.write(d.P/'closure.json',dict(status='pass',commands=commands,production_changes=0,processes=3,operations=18,archives=63,wires=18,timed_operations=0,performance_claim='none',quality='fresh standalone gates; source-identical inherited production gates'))
elif command=='seal':
 verify();assert not d.TARGET.exists() and not d.SCRATCH.exists()
 names=INDEXES+[REPORT]+[str(p.relative_to(d.ROOT)) for p in sorted(d.P.rglob('*')) if p.is_file() and p.name!='seal.json']
 d.write(d.P/'seal.json',dict(base=d.read(d.P/'origin.json')['base'],files={n:d.sha(d.ROOT/n) for n in names}))
elif command=='verify':
 verify();seal=d.read(d.P/'seal.json');assert all(d.sha(d.ROOT/n)==h for n,h in seal['files'].items())
 assert not d.TARGET.exists() and not d.SCRATCH.exists()
 if d.output(['git','rev-parse','HEAD'])!=seal['base']:
  assert d.output(['git','rev-parse','HEAD^'])==seal['base']
  changed=set(d.output(['git','diff-tree','--no-commit-id','--name-only','-r','HEAD']).splitlines())
  assert changed==set(seal['files'])|{str((d.P/'seal.json').relative_to(d.ROOT))}
else:raise AssertionError(command)
print(command+' PASS',flush=True)
