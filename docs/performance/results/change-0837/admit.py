"""Freeze exact capture dependencies and the original baseline-produced fixture."""
import hashlib,sys,time,zipfile
import driver as d
import analyze as a

COMMON=['plan.json','origin.json','host.json','driver.py','capture.py','analyze.py','audit.py','admit.py','decide.py','quality.py','probe/Cargo.toml','probe/Cargo.lock','probe/src/main.rs','build-baseline.json','build-candidate.json','freeze-baseline-v2.json','freeze-candidate.json','quality.json','supplemental-quality.json','lock-alignment.json','baseline-preflight.json','preflight-baseline-fresh-xml.json','fixtures/fixture.xml.zip','fixtures/fixture-manifest.json','fixtures/fixture.xml.zip.sha256','commands/prepare-fixtures/receipt.json','commands/prepare-fixtures/output.log']
def required(kind):
 assert kind in ['qualification','native']
 return COMMON+(['qualification.json','qualification-analysis.json','qualification-admission.json'] if kind=='native' else [])
def destination(kind):return d.P/('capture-admission.json' if kind=='native' else 'qualification-admission.json')
def fixture():
 manifest=d.read(d.P/'fixtures/fixture-manifest.json');record=d.read(d.P/'commands/prepare-fixtures/output.log');preflight=d.read(d.P/'preflight-baseline-fresh-xml.json')
 path=d.P/'fixtures/fixture.xml.zip';raw=d.sha(path)
 assert manifest['schema']=='fresh-publication-fixture-v1' and manifest['case']=='fresh-xml' and manifest['fixture']==path.name
 assert manifest['selected_xml_part']=='/custom/data.xml'
 assert raw==manifest['raw_fixture_sha256']==record['raw_fixture_sha256']==preflight['output_sha256']
 assert path.stat().st_size==manifest['fixture_bytes']==record['fixture_bytes']==preflight['output_bytes']
 assert (d.P/'fixtures/fixture.xml.zip.sha256').read_text()==f'{raw}  {path.name}\n'
 for key in ['input_semantic_sha256','member_canonical_sha256']:assert manifest[key]==record[key]==preflight[key]
 with zipfile.ZipFile(path) as archive:
  facts=[dict(name=n,bytes=len(archive.read(n)),sha256=hashlib.sha256(archive.read(n)).hexdigest()) for n in sorted(archive.namelist())]
 assert facts==preflight['members']
 receipt=d.read(d.P/'commands/prepare-fixtures/receipt.json');a.descriptor(receipt['log']);assert receipt['exit_code']==0
 assert receipt['argv']==['taskset','-c','12',d.read(d.P/'build-baseline.json')['binary']['path'],'prepare',str(d.P/'fixtures')]
 return dict(raw_sha256=raw,members=len(facts),bytes=path.stat().st_size)
def verify(kind):
 admission=d.read(destination(kind));assert admission['status']=='pass'
 assert set(admission['inputs'])==set(required(kind))
 for name,binding in admission['inputs'].items():assert d.desc(d.P/name)==binding,name
 assert d.read(d.P/'origin.json')['base']==d.read(d.P/'plan.json')['base']
 assert admission['fixture']==fixture()
 return admission
if __name__=='__main__':
 kind=sys.argv[1];d.check()
 assert d.source()==d.read(d.P/'freeze-candidate.json')['source']==d.read(d.P/'quality.json')['source']
 assert d.read(d.P/'supplemental-quality.json')['status']=='pass'
 d.write(destination(kind),dict(status='pass',created_unix=time.time(),fixture=fixture(),inputs={n:d.desc(d.P/n) for n in required(kind)}))
 verify(kind)
 print(kind,'capture admission PASS')
