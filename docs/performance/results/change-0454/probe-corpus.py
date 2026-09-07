#!/usr/bin/env python3
"""Enumerate unmodified local image slides and record real self-pair admission."""
import collections,hashlib,json,posixpath,subprocess,xml.etree.ElementTree as ET,zipfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
P='http://schemas.openxmlformats.org/presentationml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
rows=[];errors=[]
files=sorted(set(Path('test-data').rglob('*.pptx')) | set(Path('3rdparty/poi/test-data/slideshow').glob('*.pptx')) | set(Path('3rdparty/libreoffice-core/sd/qa/unit/data').rglob('*.pptx')))
for p in files:
 try:
  with zipfile.ZipFile(p) as z:
   rels={r.attrib['Id']:r.attrib['Target'] for r in ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))}
   main=ET.fromstring(z.read('ppt/presentation.xml'))
   selected=[]
   for i,s in enumerate(main.findall(f'{{{P}}}sldIdLst/{{{P}}}sldId')):
    target=rels[s.attrib[f'{{{R}}}id']];name=target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/'+target)
    slide=ET.fromstring(z.read(name));pictures=slide.findall(f'{{{P}}}cSld/{{{P}}}spTree/{{{P}}}pic')
    if pictures:selected.append((i,name,len(pictures)))
  if not selected:continue
  digest=hashlib.sha256(p.read_bytes()).hexdigest()
  for i,name,count in selected:
   cmd=['target/release/examples/native_cross_copy_probe_0454',str(p),str(i),str(p),str(i)]
   r=subprocess.run(cmd,capture_output=True,text=True)
   rows.append({'path':str(p),'sha256':digest,'bytes':p.stat().st_size,'slide':i,'slide_member':name,'pictures':count,'argv':cmd,'exit_code':r.returncode,'stdout':r.stdout,'stderr':r.stderr})
 except (OSError,KeyError,zipfile.BadZipFile,ET.ParseError,RuntimeError) as e:errors.append({'path':str(p),'error':str(e)})
result={'scope':'Local tracked and ignored unmodified files; image-bearing direct picture slides only. Same archive independently reopened for source/destination, no input rewrite. No native application or independent-pair claim.','files_inspected':len(files),'rows':rows,'scan_errors':errors,'outcomes':dict(collections.Counter((r['stdout'] or r['stderr']).strip() for r in rows))}
with (ROOT/'baseline-native-inventory.json').open('x') as f:f.write(json.dumps(result,indent=2)+'\n')
print(json.dumps({'files':len(files),'image_slides':len(rows),'errors':len(errors),'outcomes':result['outcomes']},indent=2))
