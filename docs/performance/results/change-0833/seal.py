"""Seal only owned 0833 files; verify exact committed path/blob custody."""
import subprocess
import sys
import driver as d

P=d.P
DOCS=['docs/performance/0833-filesystem-cold-qualification.md',
      *['docs/performance/'+n+'.md' for n in ['BASELINE','CRUD_COVERAGE','GOAL_AUDIT','HOTSPOTS','REPORT']]]

def paths():
    names=set(DOCS)
    for p in P.rglob('*'):
        assert not p.is_symlink(),p
        if p.is_file() and p.name!='seal.json':
            assert '__pycache__' not in p.parts and p.suffix!='.pyc'
            names.add(str(p.relative_to(d.ROOT)))
    return sorted(names)

def verify(committed=False):
    seal=d.read(P/'seal.json')
    assert seal['base']==d.BASE
    assert seal['files']=={n:d.sha(d.ROOT/n) for n in paths()}
    assert all(d.sha(d.ROOT/k)==v for k,v in d.read(P/'prepare.json')['unrelated'].items())
    assert not d.TARGET.exists() and not d.SCRATCH.exists()
    if committed:
        head=d.output(['git','rev-parse','HEAD'])
        assert d.output(['git','rev-parse','HEAD^'])==d.BASE
        changed=d.output(['git','diff-tree','--no-commit-id','--name-only','-r',head]).splitlines()
        expected=set(seal['files'])|{str((P/'seal.json').relative_to(d.ROOT))}
        assert set(changed)==expected
        import hashlib
        for name in expected:
            blob=subprocess.check_output(['git','show',f'{head}:{name}'],cwd=d.ROOT)
            assert hashlib.sha256(blob).hexdigest()==d.sha(d.ROOT/name)
    print('Seal PASS:',len(seal['files'])+1,'owned paths')

if __name__=='__main__':
    if sys.argv[1:]==['create']:
        assert d.output(['git','rev-parse','HEAD'])==d.BASE
        d.write(P/'seal.json',dict(base=d.BASE,files={n:d.sha(d.ROOT/n) for n in paths()},
                disposition='qualification_failed',formal_reports=0,performance_claim='none'))
        verify()
    else:
        assert sys.argv[1:] in (['verify'],['verify','--committed'])
        verify('--committed' in sys.argv)
