"""Build one immutable current-source diagnostic executable."""
import driver as d
import shutil,tomllib,json

d.check();source=d.source();assert source==d.read(d.P/'source.json')['source']
manifest=d.P/'probe/Cargo.toml';lock=manifest.with_name('Cargo.lock')
assert d.read(d.P/'lock-alignment.json')['status']=='pass'
assert d.desc(lock)==d.read(d.P/'lock-alignment.json')['probe']
assert d.run('probe-v2-format',['cargo','fmt','--manifest-path',str(manifest)])==0
for name,args in [('fmt',['cargo','fmt','--manifest-path',str(manifest),'--check']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',str(manifest),'--all-targets','--','-D','warnings']),('test',['cargo','test','--offline','--locked','--manifest-path',str(manifest)])]:
 assert d.run('probe-v2-'+name,args)==0
freeze=dict(source=source,probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()},driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build_v2.py'),plan=d.desc(d.P/'plan.json'))
d.write(d.P/'freeze.json',freeze)
assert d.run('probe-v2-release',['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)])==0
path=d.TARGET/'retained/opc-preservation-probe';path.parent.mkdir();shutil.copy2(d.TARGET/'release/opc-preservation-probe',path)
d.write(d.P/'build.json',dict(binary=d.desc(path),freeze=d.desc(d.P/'freeze.json')))
assert d.source()==source
print('Diagnostic build PASS')
