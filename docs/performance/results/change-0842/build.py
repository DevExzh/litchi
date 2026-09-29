"""Build equal-path retained executables from source and root-aligned lock."""
import driver as d
import shutil,sys,tomllib
leg=sys.argv[1];prefix=sys.argv[2] if len(sys.argv)>2 else leg;assert leg in ['before','after'];d.check()
manifest=d.P/'probe/Cargo.toml'
rootlock=tomllib.loads((d.ROOT/'Cargo.lock').read_text())
probelock=tomllib.loads((d.P/'probe/Cargo.lock').read_text())
reg=lambda x:{(v['name'],v['version']):(v.get('source'),v.get('checksum')) for v in x['package'] if 'source' in v}
r,p=reg(rootlock),reg(probelock);assert all(k in r and r[k]==v for k,v in p.items())
assert [k for k in p if k[0]=='flate2']==[('flate2','1.1.10')]
assert [k for k in p if k[0]=='zlib-rs']==[('zlib-rs','0.6.7')]
source=d.source();probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/f'freeze-{prefix}.json',dict(source=source,probe=probe,driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build.py'),plan=d.desc(d.P/'plan.json'),registry_packages=len(p)))
for name,args in [('fmt',['cargo','fmt','--manifest-path',str(manifest),'--check']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',str(manifest),'--all-targets','--','-D','warnings']),('test',['cargo','test','--offline','--locked','--manifest-path',str(manifest)]),('tree',['cargo','tree','--offline','--locked','--manifest-path',str(manifest),'-e','features','-i','zlib-rs']),('release',['cargo','build','--offline','--locked','--release','--manifest-path',str(manifest)])]:
 assert d.run(f'{prefix}-probe-{name}',args)==0
assert source==d.source()
assert probe=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
binary=d.TARGET/'retained'/('leg-a' if leg=='before' else 'leg-b')/'compression-probe'
binary.parent.mkdir(parents=True,exist_ok=False);shutil.copy2(d.TARGET/'release/compression-probe',binary)
codec=binary.with_name('codec-qualify');shutil.copy2(d.TARGET/'release/codec-qualify',codec)
d.write(d.P/f'build-{leg}.json',dict(binary=d.desc(binary),codec=d.desc(codec),freeze=d.desc(d.P/f'freeze-{prefix}.json'),receipt=d.desc(d.P/f'commands/{prefix}-probe-release/receipt.json')))
print('retained',leg,flush=True)
