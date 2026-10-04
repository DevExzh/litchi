"""Remove only the validated cherry-pick worktree and its build directory."""
import json
from pathlib import Path
import shutil
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=ROOT.parent/'litchi-worktrees/0913-iwa-cherry-resolution'
TARGET=ROOT/'target/litchi-0913-iwa-build'
def main():
    assert json.loads((P/'iwa-quality.json').read_text())['status']=='pass'
    assert json.loads((P/'iwa-resolution-after.json').read_text())['status']=='pass'
    assert not (P/'iwa-cleanup.json').exists()
    for d in (SOURCE,TARGET):assert not d.is_symlink() and (d/'.owned-0913-iwa').read_text()=='iwa-cherry-resolution-0913\n'
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit():continue
        try: cwd=(proc/'cwd').resolve(strict=True);exe=(proc/'exe').resolve(strict=True)
        except (OSError,RuntimeError):continue
        assert not any(cwd.is_relative_to(d) or exe.is_relative_to(d) for d in (SOURCE,TARGET)),str(proc)
    source=json.loads((P/'iwa-quality-source-v4.json').read_text())
    for row in source['files']:assert (SOURCE/row['path']).read_bytes()==(ROOT/row['path']).read_bytes(),row['path']
    changes=subprocess.check_output(['git','status','--porcelain','--untracked-files=all'],cwd=SOURCE).decode().splitlines()
    expected={row['path'] for row in source['files']}|{'.owned-0913-iwa'}
    assert {line[3:] for line in changes}==expected,changes
    removed=[]
    for d in (SOURCE,TARGET):
        files=[f for f in d.rglob('*') if f.is_file() and not f.is_symlink()]
        removed.append({'path':str(d),'files':len(files),'logical_bytes':sum(f.stat().st_size for f in files)})
    subprocess.run(['git','worktree','remove','--force',str(SOURCE)],cwd=ROOT,check=True)
    shutil.rmtree(TARGET)
    assert not SOURCE.exists() and not TARGET.exists()
    result={'status':'pass','worktrees_created':1,'worktrees_removed':1,'removed':removed}
    with (P/'iwa-cleanup.json').open('x') as out:json.dump(result,out,indent=2);out.write('\n')
    print(json.dumps(result))
if __name__=='__main__':main()
