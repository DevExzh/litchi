"""Admit remaining include-time XLSX inputs, retaining the local public fixture."""
import hashlib
import json
from pathlib import Path
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
SOURCE=Path('/tmp/litchi-0913-src')
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'

def main():
    rows=[]
    names=['docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json','3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx']
    for i,name in enumerate(names):
        data=subprocess.check_output(['git','show',PARENT+':'+name],cwd=ROOT) if i==0 else (ROOT/name).read_bytes()
        dst=SOURCE/name;assert not dst.exists();dst.parent.mkdir(parents=True,exist_ok=True);dst.write_bytes(data)
        row={'path':name,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'origin':'git' if i==0 else 'existing local LibreOffice test fixture'}
        if i==1:
            retained=P/'quality-inputs'/Path(name).name;retained.parent.mkdir();retained.write_bytes(data);row['retained']=str(retained.relative_to(ROOT))
        rows.append(row)
    with (P/'more-fixtures.json').open('x') as out:json.dump({'status':'pass','parent':PARENT,'added':rows},out,indent=2);out.write('\n')
    print('Added',len(rows),'remaining XLSX include-time inputs')

if __name__=='__main__':main()
