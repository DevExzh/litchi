"""Remove only this batch's two owned build trees after exact identity checks."""
import json,pathlib,shutil
from contract import P,ROOT,read,sha,guard
if __name__=='__main__':
    guard();assert read('analysis.json')['status']=='passed';assert read('audit.json')['status']=='passed'
    rows=[]
    for suffix,b in [('',read('build.json')[0]),('-fp',read('build-fp.json'))]:
        target=ROOT.parent/('litchi-target-0740'+suffix);exe=pathlib.Path(b['binary'])
        assert target.is_dir() and not target.is_symlink() and exe.parent==target/'release'
        assert sha(exe)==b['binary_sha256'] and exe.stat().st_size==b['binary_bytes']
        rows.append({'removed':str(target),'binary':str(exe),'sha256':sha(exe),'bytes':exe.stat().st_size})
        cachebase=pathlib.Path.home()/'.debug'/str(target).lstrip('/')
        assert cachebase.is_dir() and not cachebase.is_symlink()
        elfs=list(cachebase.rglob('elf'));assert len(elfs)==1
        cached=elfs[0];assert sha(cached)==b['binary_sha256']
        assert all(f.name in {'elf','probes'} for f in cachebase.rglob('*') if f.is_file())
        buildid=cached.parent.name
        link=pathlib.Path.home()/'.debug'/'.build-id'/buildid[:2]/buildid[2:]
        removed_link=None
        if link.is_symlink() and link.resolve().is_relative_to(cachebase):
            removed_link=str(link);link.unlink()
        rows[-1]['perf_cache']={'removed':str(cachebase),'elf_sha256':sha(cached),'removed_buildid_link':removed_link}
        shutil.rmtree(cachebase)
        shutil.rmtree(target);assert not target.exists()
    cache=P/'__pycache__'
    if cache.exists():shutil.rmtree(cache)
    guard();(P/'cleanup.json').write_text(json.dumps({'removed':rows},indent=2)+'\n')
    print('PASS exact-identity guarded cleanup of two owned build trees')
