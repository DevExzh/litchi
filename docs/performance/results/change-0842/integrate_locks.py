"""Refresh only feature dependency edges in standalone live tool locks."""
import driver as d
import shutil,tomllib,json

d.check();assert d.source()==d.read(d.P/'source-after.json')['source']
assert d.read(d.P/'quality-after.json')['status']=='pass'
changes=[]
for tool in ['perf-execution','native-resave']:
 name=f'tools/{tool}/Cargo.lock';path=d.ROOT/name
 archived=d.P/'source-before'/name;archived.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(path,archived)
 before=tomllib.loads(path.read_text())
 assert d.run(f'{tool}-feature-lock',['cargo','metadata','--offline','--manifest-path',str(path.with_name('Cargo.toml')),'--format-version','1'])==0
 after=tomllib.loads(path.read_text())
 # Preserve duplicate package names: compare every non-flate2 record exactly.
 assert {k:v for k,v in before.items() if k!='package'}=={k:v for k,v in after.items() if k!='package'}
 canonical=lambda rows:sorted(json.dumps(x,sort_keys=True) for x in rows)
 assert canonical(x for x in before['package'] if x['name']!='flate2')==canonical(x for x in after['package'] if x['name']!='flate2')
 a=[x for x in before['package'] if x['name']=='flate2'];b=[x for x in after['package'] if x['name']=='flate2']
 assert len(a)==len(b)==1
 a,b=a[0],b[0]
 assert a['version']==b['version']=='1.1.10'
 assert {k:v for k,v in a.items() if k!='dependencies'}=={k:v for k,v in b.items() if k!='dependencies'}
 assert a['dependencies']==['zlib-rs'] and b['dependencies']==['crc32fast','zlib-rs']
 target=d.P/'source-final'/name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(path,target)
 changes.append(dict(path=name,before=d.desc(archived),after=d.desc(path),added_dependency='crc32fast',package_versions_unchanged=True))
receipts=[]
for tool in ['perf-baseline','perf-execution','native-resave']:
 label=tool+'-locked-metadata'
 assert d.run(label,['cargo','metadata','--offline','--locked','--manifest-path',str(d.ROOT/'tools'/tool/'Cargo.toml'),'--format-version','1'])==0
 receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
source=d.source();measured=d.read(d.P/'source-after.json')['source']
delta=sorted(n for n in set(source)|set(measured) if source.get(n)!=measured.get(n))
assert delta==['tools/native-resave/Cargo.lock','tools/perf-execution/Cargo.lock']
d.write(d.P/'source-final.json',dict(source=source,delta_from_measured_after=delta,note='Only standalone tool lock feature edges changed after the measured candidate build and owner gates. Neither nested lock is an input to the isolated probe or root workspace builds. Registry versions/checksums remain unchanged.'))
d.write(d.P/'nested-lock-checks.json',dict(status='pass',changes=changes,locked_metadata_receipts=receipts))
print('All three committed standalone tool locks pass locked metadata; only two additional feature edges')
