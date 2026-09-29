"""Build and qualify one frozen diagnostic leg; root invokes serially."""
import driver as d
import shutil,sys
leg=sys.argv[1];assert leg in ['before','after']
manifest=str(d.P/'probe/Cargo.toml');d.check()
source=d.source();probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/f'freeze-{leg}.json',dict(source=source,probe=probe,driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build.py'),plan=d.desc(d.P/'plan.json')))
checks=[('fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),('check',['cargo','check','--offline','--locked','--manifest-path',manifest,'--all-targets','--all-features']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',manifest,'--all-targets','--all-features','--','-D','warnings']),('doc',['cargo','doc','--offline','--locked','--manifest-path',manifest,'--all-features','--no-deps']),('test-default',['cargo','test','--offline','--locked','--manifest-path',manifest]),('test-all',['cargo','test','--offline','--locked','--manifest-path',manifest,'--all-features','--','--test-threads=1'])]
receipts=[]
for name,argv in checks:
 label=f'{leg}-probe-{name}';code=d.run(label,argv);receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
 if code:sys.exit(code)
d.write(d.P/f'probe-quality-{leg}.json',dict(status='pass',receipts=receipts))
assert d.run('build-'+leg,['cargo','build','--offline','--locked','--release','--manifest-path',manifest])==0
binary=d.TARGET/'retained'/('leg-a' if leg=='before' else 'leg-b')/'ppt-profile-probe'
binary.parent.mkdir(parents=True,exist_ok=False);shutil.copy2(d.TARGET/'release/ppt-profile-probe',binary)
assert d.source()==source and all(d.sha(d.P/n)==h for n,h in probe.items())
d.write(d.P/f'build-{leg}.json',dict(binary=d.desc(binary),freeze=d.desc(d.P/f'freeze-{leg}.json'),receipt=d.desc(d.P/'commands'/f'build-{leg}'/'receipt.json')))
print('Built',leg,flush=True)
