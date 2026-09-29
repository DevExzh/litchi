"""Independent stdlib ZIP/XML replay; runnable after temporary roots are removed."""
import hashlib,json,struct,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
P=Path(__file__).resolve().parent
H=lambda b:hashlib.sha256(b).hexdigest()
read=lambda p:json.loads(p.read_text())
MAIN='word/document.xml'
def facts(p):
 raw=p.read_bytes();logical={};compressed={};local={}
 with zipfile.ZipFile(p) as z:
  names=sorted(z.namelist());assert len(names)==len(set(names))
  for n in names:
   i=z.getinfo(n);logical[n]=z.read(n)
   off=i.header_offset;header=raw[off:off+30];assert header[:4]==b'PK\x03\x04'
   nl,xl=struct.unpack_from('<HH',header,26);start=off+30+nl+xl
   compressed[n]=raw[start:start+i.compress_size]
   local[n]=raw[off:start+i.compress_size]
 canonical=hashlib.sha256()
 def tag(b):canonical.update(struct.pack('<Q',len(b)));canonical.update(b)
 tag(b'docx-diagnostic-members-v1');canonical.update(struct.pack('<Q',len(names)))
 members=[]
 for n in names:
  tag(n.encode());tag(logical[n]);members.append(dict(name=n,bytes=len(logical[n]),sha256=H(logical[n])))
 return dict(bytes=len(raw),raw_sha256=H(raw),member_canonical_sha256=canonical.hexdigest(),members=members),logical,compressed,local
def paragraphs(logical):
 root=ET.fromstring(logical[MAIN]);ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
 return [''.join(p.itertext()) for p in root.findall('./w:body/w:p',ns)]
def check_record(folder,r):
 f,l,c,s=facts(folder/r['archive'])
 for k,v in f.items():assert r[k]==v,(r['archive'],k)
 assert r['logical_equal'] is True
 return l,c,s
rows=[];pids=[];inventories=[]
for folder in sorted((P/'runs').iterdir()):
 report=read(folder/'report.json');pids.append(report['pid']);sources={}
 assert len(report['sources'])==3 and len(report['operations'])==6
 for r in report['sources']:
  label=r['archive'][7:-4];sources[label]=check_record(folder,r)
 assert set(sources)=={'original-retained','borrowed-regenerated','current-owned-staged'}
 original=sources['original-retained'][0]
 assert (folder/'source-original-retained.zip').read_bytes()==(P/'fixtures/fresh-source.zip').read_bytes()
 for logical,_,_ in sources.values():assert logical==original and paragraphs(logical)==['one','two']
 seen=set()
 for op in report['operations']:
  label=op['source_variant'];kind='copy' if op['case']=='plain-paragraph-copy-0-to-2' else 'removal'
  assert op['case'] in ['plain-paragraph-copy-0-to-2','plain-paragraph-removal-0']
  assert (label,kind) not in seen;seen.add((label,kind))
  source,compressed,local=sources[label]
  forward,fc,fl=check_record(folder,op['forward']);durable,dc,dl=check_record(folder,op['durable_inverse']);immediate,ic,il=check_record(folder,op['immediate_inverse'])
  assert paragraphs(forward)==(['one','two','one'] if kind=='copy' else ['two'])
  assert set(forward)==set(source) and durable==immediate==source
  for n in source:
   if n!=MAIN:assert forward[n]==source[n] and fl[n]==dl[n]==local[n]
  source_bytes=(folder/f'source-{label}.zip').read_bytes()
  assert (folder/op['immediate_inverse']['archive']).read_bytes()==source_bytes
  exact=(folder/op['durable_inverse']['archive']).read_bytes()==source_bytes
  assert exact==op['durable_inverse_exact']
  assert all(op[k] is True for k in ['logical_equal','immediate_inverse_exact','stale_refused','stale_sink_empty'])
  wire=op['durable_wire'];b=(folder/wire['path']).read_bytes()
  assert len(b)==wire['bytes'] and H(b)==wire['sha256']
  assert b[:8]==(b'LDXPCPY\0' if kind=='copy' else b'LDXPREM\0') and b[8]==1
  rows.append(dict(run=folder.name,source=label,operation=kind,durable_exact=exact,changed_compressed_members=[n for n in source if compressed[n]!=dc[n]],source_main_compressed_bytes=len(compressed[MAIN]),restored_main_compressed_bytes=len(dc[MAIN])))
 assert len(seen)==6
 inventories.append({p.name:H(p.read_bytes()) for p in folder.iterdir() if p.name!='report.json'})
assert len(pids)==len(set(pids))==3
assert inventories[0]==inventories[1]==inventories[2]
assert len(inventories[0])==27
for x in read(P/'capture.json')['artifacts']:
 p=P/'runs'/Path(x['path']).parent.name/Path(x['path']).name
 assert p.stat().st_size==x['bytes'] and H(p.read_bytes())==x['sha256']
result=dict(status='pass',processes=3,operations=len(rows),archives=63,wires=18,distinct_pids=pids,cross_process_bytes_identical=True,rows=rows,scope='ZIP/XML independently replayed; canonical wire re-encoding and stale refusal checked by Rust probe')
expected=P/'analysis.json'
if expected.exists():assert read(expected)==result
else:expected.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print(json.dumps(result,indent=2))
