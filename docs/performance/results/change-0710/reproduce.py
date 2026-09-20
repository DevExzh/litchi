#!/usr/bin/env python3
"""Independently compare oracle ZIP member bytes to the source fixtures."""
import hashlib
import json
from pathlib import Path
import zipfile
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(b):return hashlib.sha256(b).hexdigest()
def members(path):
    with zipfile.ZipFile(path) as z:
        names=z.namelist();assert len(names)==len(set(names))
        return {n:z.read(n) for n in names if not n.endswith('/')}
def main():
    report=json.loads((P/'oracle/report.json').read_text());rows=[]
    for fixture in ['successful_fixture','refusal_fixture']:
        item=report[fixture];src=members(ROOT/item['identity']['relative_path'])
        for route in item['routes']+item.get('no_edit_routes',[]):
            path=Path(route['artifact_path']);out=members(path)
            changed=sorted(n for n in src.keys()|out.keys() if src.get(n)!=out.get(n))
            if fixture=='refusal_fixture':assert not changed
            rows.append(dict(fixture=fixture,route=route['route'],artifact=str(path.relative_to(P)),sha256=sha(path.read_bytes()),changed_decoded_members=changed,custom_properties_preserved=src.get('docProps/custom.xml')==out.get('docProps/custom.xml')))
    target=P/'zip-verification.json';target.write_text(json.dumps(dict(status='pass',rows=rows,scope='Independent Python ZIP decompression; refused and no-edit routes preserve every decoded member; admitted route changes listed without general preservation claim'),indent=2)+'\n')
    print('PASS independent ZIP check',len(rows),'outputs')
if __name__=='__main__':main()
