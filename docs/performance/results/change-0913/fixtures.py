"""Temporarily admit archived test fixtures, then restore benchmark inventory."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile
import tempfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/tmp/litchi-0913-src')
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'
SUPPLEMENTAL='crates/soapberry-zip/tests/owned_direct_adversarial.rs'
SUPPLEMENTAL_REVISION=P/'supplemental-owned-direct-adversarial-v2.json'

def sha(data): return hashlib.sha256(data).hexdigest()

def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('mode',choices=['add','remove']); args=parser.parse_args()
    assert (SOURCE/'.owned-0913').read_text()=='owned-direct-read-0913\n'
    if args.mode=='add':
        archive=subprocess.check_output(['git','archive',PARENT,'--','test-data'],cwd=ROOT)
        added=[]
        with tempfile.TemporaryDirectory(prefix='litchi-0913-fixtures-') as temp:
            with tarfile.open(fileobj=io.BytesIO(archive)) as stream: stream.extractall(temp,filter='data')
            for f in sorted(Path(temp).rglob('*')):
                if not f.is_file(): continue
                name=str(f.relative_to(temp)); dst=SOURCE/name; data=f.read_bytes()
                if dst.exists():
                    assert dst.read_bytes()==data, name
                    continue
                dst.parent.mkdir(parents=True,exist_ok=True)
                dst.write_bytes(data)
                added.append({'path':name,'bytes':len(data),'sha256':sha(data)})
        name='crates/soapberry-zip/tests/owned_direct_adversarial.rs'
        data=(ROOT/name).read_bytes(); dst=SOURCE/name
        assert not dst.exists(); dst.write_bytes(data)
        added.append({'path':name,'bytes':len(data),'sha256':sha(data),'origin':'workspace supplemental test'})
        record={'status':'pass','parent':PARENT,'archive_paths':['test-data'],'archive_sha256':sha(archive),'added':added,'benchmark_production_files_changed':False}
        with (P/'fixtures.json').open('x') as f:json.dump(record,f,indent=2);f.write('\n')
        print('Added',len(added),'archived test fixtures')
    else:
        record=json.loads((P/'fixtures.json').read_text())
        extra=json.loads((P/'extra-fixtures.json').read_text())
        more=json.loads((P/'more-fixtures.json').read_text())
        runtime=json.loads((P/'runtime-fixtures.json').read_text())
        revision=json.loads(SUPPLEMENTAL_REVISION.read_text())
        assert revision['status']=='pass' and revision['kind']=='supplemental-test-revision' and revision['revision']==2
        assert revision['path']==SUPPLEMENTAL
        previous=revision['previous']; current=revision['current']
        supplemental_rows=[row for row in record['added'] if row['path']==SUPPLEMENTAL]
        assert len(supplemental_rows)==1
        old_row=supplemental_rows[0]
        assert previous['path']=='docs/performance/results/change-0913/quality-inputs/owned_direct_adversarial-v1.rs'
        assert previous['bytes']==old_row['bytes']==9494 and previous['sha256']==old_row['sha256']
        assert current['path']==SUPPLEMENTAL
        for row in record['added']+extra['added']+more['added']+runtime['added']:
            if row['path']==SUPPLEMENTAL:
                f=SOURCE/current['path']
                assert f.stat().st_size==current['bytes'] and sha(f.read_bytes())==current['sha256']
            else:
                f=SOURCE/row['path']
                assert f.stat().st_size==row['bytes'] and sha(f.read_bytes())==row['sha256']
            f.unlink()
        for d in sorted((SOURCE/'test-data').rglob('*'),key=lambda p:len(p.parts),reverse=True):
            if d.is_dir() and not any(d.iterdir()):d.rmdir()
        lock=json.loads((P/'quality-lock-v3/command.json').read_text())
        assert sha((SOURCE/'Cargo.lock').read_bytes())==lock['after']['sha256']
        (SOURCE/'Cargo.lock').write_bytes((P/'quality-lock/before-Cargo.lock').read_bytes())
        frozen=json.loads((P/'source-v1.json').read_text())
        actual=[{'path':str(f.relative_to(SOURCE)),'bytes':f.stat().st_size,'sha256':sha(f.read_bytes())} for f in sorted(SOURCE.rglob('*')) if f.is_file()]
        assert actual==frozen['files']
        with (P/'fixtures-removed.json').open('x') as f:json.dump({'status':'pass','removed':len(record['added']),'extra_removed':len(extra['added']),'more_removed':len(more['added']),'runtime_removed':len(runtime['added']),'benchmark_inventory_restored':True},f,indent=2);f.write('\n')
        print('Restored frozen benchmark inventory')

if __name__=='__main__':main()
