"""Remove only marked 0913 source/build roots after successful replay."""
import hashlib
import json
from pathlib import Path
import shutil

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
ROOTS=[Path('/tmp/litchi-0913-src'),ROOT/'target/litchi-0913-build']
MARKER='owned-direct-read-0913\n'

def main():
    assert json.loads((P/'replay-before-cleanup.json').read_text())['status']=='pass'
    assert not (P/'cleanup.json').exists()
    for root in ROOTS:
        assert not root.is_symlink() and (root/'.owned-0913').read_text()==MARKER
    for proc in Path('/proc').iterdir():
        if not proc.name.isdigit(): continue
        try:
            cwd=(proc/'cwd').resolve(strict=True)
            exe=(proc/'exe').resolve(strict=True)
        except (OSError,RuntimeError): continue
        assert not any(cwd.is_relative_to(root) or exe.is_relative_to(root) for root in ROOTS), (proc,cwd,exe)
    removed=[]
    for root in ROOTS:
        files=[f for f in root.rglob('*') if f.is_file() and not f.is_symlink()]
        removed.append({'path':str(root),'files':len(files),'logical_bytes':sum(f.stat().st_size for f in files)})
        shutil.rmtree(root)
        assert not root.exists()
    result={'status':'pass','worktrees_created':0,'removed':removed,'retained_evidence':'candidate binary and capture source/output PPTX archives remain locally, gitignored'}
    with (P/'cleanup.json').open('x') as out: json.dump(result,out,indent=2);out.write('\n')
    print(json.dumps(result))

if __name__=='__main__': main()
