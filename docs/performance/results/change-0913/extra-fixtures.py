"""Admit the archived ZIP64 corpus required by existing integration tests."""
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/tmp/litchi-0913-src')
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'

def main():
    scope='docs/performance/results/change-0416/corpus'
    archive=subprocess.check_output(['git','archive',PARENT,'--',scope],cwd=ROOT)
    rows=[]
    with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
        for member in stream.getmembers():
            if not member.isfile():continue
            assert member.name.startswith(scope+'/') and '..' not in Path(member.name).parts
            data=stream.extractfile(member).read(); dst=SOURCE/member.name
            assert not dst.exists();dst.parent.mkdir(parents=True,exist_ok=True);dst.write_bytes(data)
            rows.append({'path':member.name,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()})
    with (P/'extra-fixtures.json').open('x') as out:json.dump({'status':'pass','parent':PARENT,'archive_paths':[scope],'archive_sha256':hashlib.sha256(archive).hexdigest(),'added':rows},out,indent=2);out.write('\n')
    print('Added',len(rows),'archived ZIP64 fixtures')

if __name__=='__main__':main()
