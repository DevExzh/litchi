"""Build matched immutable binaries from one frozen probe."""
import driver as d
import sys,shutil,tomllib
leg=sys.argv[1];assert leg in ['before','after'];d.check();source=d.source()
assert source==d.read(d.P/f'source-{leg}.json')['source']
manifest=d.P/'probe/Cargo.toml';lock=manifest.with_name('Cargo.lock')
if leg=='before':
 shutil.copy2(d.P.parent/'change-0844/probe/Cargo.lock',lock)
 assert d.run('before-probe-format',['cargo','fmt','--manifest-path',str(manifest)])==0
 root=tomllib.loads((d.ROOT/'Cargo.lock').read_text())['package'];probe=tomllib.loads(lock.read_text())['package']
 registry=lambda rows:{(x['name'],x['version'],x.get('source')):x.get('checksum') for x in rows if 'source' in x}
 a,b=registry(root),registry(probe);assert all(a.get(k)==v for k,v in b.items())
 d.write(d.P/'lock-alignment.json',dict(status='pass',root=d.desc(d.ROOT/'Cargo.lock'),probe=d.desc(lock),registry_packages=len(b)))
else:
 frozen=d.read(d.P/'freeze-before.json');assert all(d.sha(d.P/n)==h for n,h in frozen['probe'].items())
for name,args in [('fmt',['cargo','fmt','--manifest-path',str(manifest),'--check']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',str(manifest),'--all-targets','--','-D','warnings']),('test',['cargo','test','--offline','--locked','--manifest-path',str(manifest)])]:
 assert d.run(leg+'-probe-'+name,args)==0
freeze=dict(source=source,probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()},driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build.py'),plan=d.desc(d.P/'plan.json'))
d.write(d.P/f'freeze-{leg}.json',freeze)
assert d.run(leg+'-probe-release',['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)])==0
path=d.TARGET/'retained'/leg/'opc-preservation-probe';path.parent.mkdir(parents=True);shutil.copy2(d.TARGET/'release/opc-preservation-probe',path)
d.write(d.P/f'build-{leg}.json',dict(binary=d.desc(path),freeze=d.desc(d.P/f'freeze-{leg}.json')))
assert d.source()==source
print(leg,'build PASS')
