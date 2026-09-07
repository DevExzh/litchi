#!/usr/bin/env python3
"""Capture actual distinct-package admission without rewriting fixtures."""
import hashlib,importlib.util,json,subprocess
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3]
spec=importlib.util.spec_from_file_location('custody',ROOT/'check.py');c=importlib.util.module_from_spec(spec);spec.loader.exec_module(c)
manifest=json.loads((ROOT/'probe-build.json').read_text());binary=manifest['binary'];assert hashlib.sha256(Path(binary['path']).read_bytes()).hexdigest()==binary['sha256']
a='3rdparty/libreoffice-core/sd/qa/unit/data/smoketest.pptx';b='3rdparty/poi/test-data/slideshow/at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx'
rows=[]
for source,sp,destination,dp in [(a,0,b,i) for i in [16,17,33]]+[(b,i,a,0) for i in [16,17,33]]:
    before=c.sources();argv=[binary['path'],source,str(sp),destination,str(dp)];r=subprocess.run(argv,cwd=REPO,capture_output=True,text=True);after=c.sources();assert before==after
    def identity(name):
        p=REPO/name;return {'path':name,'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size}
    rows.append({'source':identity(source),'destination':identity(destination),'source_position':sp,'destination_position':dp,'argv':argv,'source_before':before,'source_after':after,'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
record={'scope':'Six distinct original-file pair admissions; no rewriting, no measured timing or native application claim. Typed refusal is an outcome, not a successful copy.','binary':binary,'driver_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),'rows':rows}
with (ROOT/'pair-probe.json').open('x') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps(record))
