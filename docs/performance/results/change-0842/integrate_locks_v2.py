"""Finish lock integration after detecting pre-existing native-resave lock drift."""
import driver as d
import json,shutil,tomllib
d.check();assert d.read(d.P/'quality-after.json')['status']=='pass'
name='tools/native-resave/Cargo.lock';path=d.ROOT/name
candidate=path.read_bytes();manifests=['Cargo.toml','crates/soapberry-zip/Cargo.toml'];saved={n:(d.ROOT/n).read_bytes() for n in manifests}
try:
 for n in manifests:(d.ROOT/n).write_bytes((d.P/'source-before'/n).read_bytes())
 path.write_bytes((d.P/'source-before'/name).read_bytes())
 assert d.run('native-resave-baseline-lock-control',['cargo','metadata','--offline','--manifest-path',str(path.with_name('Cargo.toml')),'--format-version','1'])==0
 normalized=path.read_bytes()
finally:
 for n,v in saved.items():(d.ROOT/n).write_bytes(v)
 path.write_bytes(candidate)
archive=d.P/'lock-snapshots/native-resave-baseline-normalized.lock';archive.write_bytes(normalized)
a=tomllib.loads(normalized.decode());b=tomllib.loads(candidate.decode())
canonical=lambda rows:sorted(json.dumps(x,sort_keys=True) for x in rows)
assert {k:v for k,v in a.items() if k!='package'}=={k:v for k,v in b.items() if k!='package'}
assert canonical(x for x in a['package'] if x['name']!='flate2')==canonical(x for x in b['package'] if x['name']!='flate2')
af=[x for x in a['package'] if x['name']=='flate2'];bf=[x for x in b['package'] if x['name']=='flate2'];assert len(af)==len(bf)==1
af,bf=af[0],bf[0]
assert {k:v for k,v in af.items() if k!='dependencies'}=={k:v for k,v in bf.items() if k!='dependencies'}
assert af['dependencies']==['zlib-rs'] and bf['dependencies']==['crc32fast','zlib-rs']
old=tomllib.loads((d.P/'source-before'/name).read_text())
key=lambda x:(x['name'],x['version'],x.get('source',''))
o={key(x):x for x in old['package']};n={key(x):x for x in a['package']};assert len(o)==len(old['package']) and len(n)==len(a['package'])
added=[n[k] for k in n.keys()-o.keys()];removed=[o[k] for k in o.keys()-n.keys()];changed=[dict(before=o[k],after=n[k]) for k in sorted(o.keys()&n.keys()) if o[k]!=n[k]]
assert sorted(x['name'] for x in added)==['same-file','winapi-util'] and not removed
assert sorted(x['before']['name'] for x in changed)==['litchi-doc-crypto','litchi-doc-write','litchi-opc','litchi-ppt-write']
assert all('source' not in x['before'] for x in changed)
d.write(d.P/'native-resave-lock-drift.json',dict(status='confirmed-preexisting-lock-drift',failed_guard='integrate_locks.py rejected non-flate2 changes before native timing',baseline_normalized=d.desc(archive),baseline_receipt=d.desc(d.P/'commands/native-resave-baseline-lock-control/receipt.json'),baseline_manifests={n:d.sha(d.P/'source-before'/n) for n in manifests},added=added,removed=removed,changed=changed,interpretation='Cargo refreshes stale local dependency edges and adds their two missing transitive packages even with original compression declarations. The candidate adds only crc32fast to flate2 relative to this normalized baseline. No existing registry version or checksum changes.'))
changes=[]
for tool in ['perf-execution','native-resave']:
 name=f'tools/{tool}/Cargo.lock';path=d.ROOT/name;target=d.P/'source-final'/name;target.parent.mkdir(parents=True,exist_ok=True)
 if target.exists():assert target.read_bytes()==path.read_bytes()
 else:shutil.copyfile(path,target)
 changes.append(dict(path=name,before=d.desc(d.P/'source-before'/name),after=d.desc(path),added_dependency='crc32fast',existing_registry_versions_unchanged=True))
receipts=[]
for tool in ['perf-baseline','perf-execution','native-resave']:
 label=tool+'-locked-metadata'
 assert d.run(label,['cargo','metadata','--offline','--locked','--manifest-path',str(d.ROOT/'tools'/tool/'Cargo.toml'),'--format-version','1'])==0
 receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
source=d.source();measured=d.read(d.P/'source-after.json')['source'];delta=sorted(n for n in set(source)|set(measured) if source.get(n)!=measured.get(n))
assert delta==['tools/native-resave/Cargo.lock','tools/perf-execution/Cargo.lock']
d.write(d.P/'source-final.json',dict(source=source,delta_from_measured_after=delta,note='Standalone tool locks are not probe/root build inputs. perf-execution gains only the flate2 feature edge. native-resave also refreshes baseline-confirmed stale local dependency edges and adds same-file/winapi-util, with no existing registry version/checksum changes. See native-resave-lock-drift.json.'))
d.write(d.P/'nested-lock-checks.json',dict(status='pass',changes=changes,baseline_lock_drift=d.desc(d.P/'native-resave-lock-drift.json'),locked_metadata_receipts=receipts))
print('All three committed tool locks pass locked metadata; native-resave drift independently reproduced on baseline')
