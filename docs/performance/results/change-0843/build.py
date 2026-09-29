"""Build one immutable current-source diagnostic executable."""
import driver as d
import shutil,tomllib,json

d.check();source=d.source();assert source==d.read(d.P/'source.json')['source']
manifest=d.P/'probe/Cargo.toml';lock=manifest.with_name('Cargo.lock')
shutil.copyfile(d.ROOT/'Cargo.lock',lock)
assert d.run('probe-lock',['cargo','metadata','--offline','--manifest-path',str(manifest),'--format-version','1'])==0
root=tomllib.loads((d.ROOT/'Cargo.lock').read_text());probe=tomllib.loads(lock.read_text())
registry=lambda v:{(x['name'],x['version'],x['source']):x['checksum'] for x in v['package'] if 'source' in x}
a,b=registry(root),registry(probe);assert all(a.get(k)==v for k,v in b.items())
d.write(d.P/'lock-alignment.json',dict(status='pass',root=d.desc(d.P/'root-Cargo.lock'),probe=d.desc(lock),registry_packages=len(b)))
for name,args in [('fmt',['cargo','fmt','--manifest-path',str(manifest),'--check']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',str(manifest),'--all-targets','--','-D','warnings']),('test',['cargo','test','--offline','--locked','--manifest-path',str(manifest)])]:
 assert d.run('probe-'+name,args)==0
freeze=dict(source=source,probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()},driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build.py'),plan=d.desc(d.P/'plan.json'))
d.write(d.P/'freeze.json',freeze)
assert d.run('probe-release',['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)])==0
path=d.TARGET/'retained/opc-preservation-probe';path.parent.mkdir();shutil.copy2(d.TARGET/'release/opc-preservation-probe',path)
d.write(d.P/'build.json',dict(binary=d.desc(path),freeze=d.desc(d.P/'freeze.json')))
assert d.source()==source
print('Diagnostic build PASS')
