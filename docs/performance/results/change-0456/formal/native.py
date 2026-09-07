#!/usr/bin/env python3
"""Check the candidate on the pinned LibreOffice QA self-pair fixture."""
import hashlib,json,subprocess,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[4]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
expected=json.loads((ROOT/'native-expected.json').read_text());build=json.loads((ROOT/'candidate-build.json').read_text());binary=build['binaries']['external'];assert sha(Path(binary['path']))==binary['sha256']
fixture=REPO/'3rdparty/libreoffice-core/sd/qa/unit/data/smoketest.pptx';assert sha(fixture)==expected['fixture_sha256']
task=Path('/tmp/litchi-goal-0456/native');task.mkdir();(ROOT/'native').mkdir()
records=[]
for provider in ['bytes','range']:
    report=task/(provider+'.json');argv=[binary['path'],str(fixture),str(report),'1','0',build['revision'],provider]
    result=subprocess.run(argv,capture_output=True,text=True);assert result.returncode==0,result.stderr
    output=report.with_suffix('.pptx');assert output.stat().st_size==expected['output_bytes'];assert sha(output)==expected['output_sha256']
    with zipfile.ZipFile(output) as z:
        assert z.testzip() is None
        assert z.namelist()==expected['output_member_order']
        for name,digest in [('ppt/slides/slide2.xml',expected['copied_slide_sha256']),('ppt/media/image1-copy1.png',expected['copied_image_sha256']),*expected['metadata_sha256'].items()]:assert hashlib.sha256(z.read(name)).hexdigest()==digest
        for name in z.namelist():
            if name.endswith(('.xml','.rels')):ET.fromstring(z.read(name))
    retained=ROOT/'native'/(provider+'.json');retained.write_bytes(report.read_bytes())
    records.append({'provider':provider,'argv':argv,'exit_code':result.returncode,'stdout':result.stdout,'stderr':result.stderr,'report':{'path':str(retained.relative_to(ROOT)),'bytes':retained.stat().st_size,'sha256':sha(retained)},'output':{'path':str(output),'bytes':output.stat().st_size,'sha256':sha(output)}})
(ROOT/'native-proof.json').write_text(json.dumps({'status':'pass','scope':'same pinned LibreOffice QA archive reopened independently as source and destination; exact 0454 output hash, ZIP CRC, member order, copied bytes and XML parsing; no native-application roundtrip','fixture_sha256':sha(fixture),'expected_sha256':sha(ROOT/'native-expected.json'),'binary':binary,'rows':records},indent=2)+'\n')
print(json.dumps({'status':'pass','providers':2}))
