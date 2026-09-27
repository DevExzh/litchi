"""Offline whole-packet custody and workspace preservation; no native runs."""
import subprocess,sys
import common as c
import analyze

result=analyze.check()
assert result==c.read(c.P/'analysis.json')
# A replay pass preserves the frozen experiment's failures, never overrides them.
assert result['acceptance']=='fail'
assert [(x['lane'],x['index']) for x in result['acceptance_failures']]==[('traces',0),('traces',1),('traces',2)]
source=c.read(c.verify(c.read(c.P/'inherited.json')['production-source.json']))['files']
for n,h in source.items():assert c.sha(c.ROOT/n)==h,n
arch=c.read(c.P/'architecture-inputs.json')
for n,h in arch.items():assert c.sha(c.ROOT/n)==h,n
if '--check-workspace' in sys.argv:
 origin=c.read(c.P/'origin.json')
 for n,h in origin['unrelated'].items():assert c.sha(c.ROOT/n)==h,n
 actual=subprocess.check_output(['git','worktree','list','--porcelain'],cwd=c.ROOT,text=True)
 assert actual.split('\n\n',1)[1]==origin['worktrees'].split('\n\n',1)[1]
if '--require-final-seal' in sys.argv:
 seal=c.read(c.P/'seal.json');files={str(p.relative_to(c.P)) for p in c.P.rglob('*') if p.is_file() and p.name!='seal.json'}
 assert files==set(seal['files'])
 for n,a in seal['files'].items():assert a['path']==n;c.verify(a)
 for a in seal['documents']:c.verify(a)
 assert seal['acceptance']==result['acceptance'] and seal['acceptance_failures']==result['acceptance_failures']
 assert not c.TARGET.exists() and not list(c.P.rglob('__pycache__'))
print(f'0790 custody/replay PASS: {len(source)} production files, {len(arch)} architecture inputs; experiment acceptance remains FAIL (three trace controls)')
