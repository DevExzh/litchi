"""Reconstruct the missing ephemeral source from retained, hash-bound evidence."""
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/tmp/litchi-0913-src')
OLD=P.parent/'change-0910'
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'
def sha(data): return hashlib.sha256(data).hexdigest()
def load(path): return json.loads(path.read_text())
def desc(path): return {'path':str(path.relative_to(ROOT)), 'bytes':path.stat().st_size,'sha256':sha(path.read_bytes())}
def write(name,data,row):
    assert len(data)==row['bytes'] and sha(data)==row['sha256'],name
    f=SOURCE/name; f.parent.mkdir(parents=True,exist_ok=True); f.write_bytes(data)
def archive(commit,paths):
    return subprocess.check_output(['git','archive',commit,'--',*paths],cwd=ROOT)
def inventory():
    return [{'path':str(f.relative_to(SOURCE)),'bytes':f.stat().st_size,'sha256':sha(f.read_bytes())} for f in sorted(SOURCE.rglob('*')) if f.is_file()]
def main():
    assert not SOURCE.exists() and not (P/'source-recovery.json').exists()
    base=load(OLD/'source-base.json'); old=load(OLD/'source-v1.json'); frozen=load(P/'source-v1.json')
    data=archive(base['commit'],base['archive_paths']); assert sha(data)==base['archive_sha256']
    SOURCE.mkdir()
    with tarfile.open(fileobj=io.BytesIO(data)) as stream: stream.extractall(SOURCE,filter='data')
    for name,row in base['lock_overlays'].items():
        data=(ROOT/row['artifact']).read_bytes(); assert sha(data)==row['sha256']; (SOURCE/name).write_bytes(data)
    (SOURCE/'.owned-0910').write_text(base['commit']+'\n')
    for manifest in (old,frozen):
        row=manifest['overlay_patch']; patch=ROOT/row['path']; assert sha(patch.read_bytes())==row['sha256']
        subprocess.run(['git','apply',str(patch)],cwd=SOURCE,check=True)
        if manifest is frozen: (SOURCE/'.owned-0913').write_text('owned-direct-read-0913\n')
        assert inventory()==manifest['files']
    for filename in ('fixtures.json','extra-fixtures.json'):
        record=load(P/filename); data=archive(PARENT,record['archive_paths']); assert sha(data)==record['archive_sha256']
        with tarfile.open(fileobj=io.BytesIO(data)) as stream:
            for row in record['added']:
                if row.get('origin')=='workspace supplemental test': continue
                write(row['path'],stream.extractfile(row['path']).read(),row)
    for filename in ('more-fixtures.json','runtime-fixtures.json'):
        for row in load(P/filename)['added']:
            data=(ROOT/row['retained']).read_bytes() if 'retained' in row else subprocess.check_output(['git','show',PARENT+':'+row['path']],cwd=ROOT)
            write(row['path'],data,row)
    revision=load(P/'supplemental-owned-direct-adversarial-v2.json'); row=revision['current']; write(row['path'],(ROOT/row['path']).read_bytes(),row)
    lock=load(P/'quality-lock-v3/command.json'); data=(P/'quality-lock-v3/Cargo.lock').read_bytes(); assert sha(data)==lock['after']['sha256']; (SOURCE/'Cargo.lock').write_bytes(data)
    result={'status':'pass','reason':'Temporary source observed absent after environment continuation; cause unknown. No deliberate cleanup receipt existed.','source_root':str(SOURCE),'baseline_files':len(old['files']),'candidate_files':len(frozen['files']),'quality_files':len(inventory()),'helper':desc(Path(__file__)),'source':desc(P/'source-v1.json'),'supplemental_revision':desc(P/'supplemental-owned-direct-adversarial-v2.json'),'production_changes':False,'candidate_binary':desc(P/'binaries/candidate')}
    with (P/'source-recovery.json').open('x') as out: json.dump(result,out,indent=2); out.write('\n')
    print(json.dumps(result))
if __name__=='__main__': main()
