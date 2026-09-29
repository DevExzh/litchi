"""Qualify and retain native and allocation binaries from identical sources."""
import driver as d
import shutil,sys
leg=sys.argv[1];quality_label=sys.argv[2] if len(sys.argv)>2 else leg;manifest=str(d.P/'probe/Cargo.toml')
d.check();source=d.source();probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/f'probe-inputs-{quality_label}.json',dict(source=source,probe=probe))
checks=[('fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),('check',['cargo','check','--offline','--locked','--manifest-path',manifest,'--all-targets','--all-features']),('clippy',['cargo','clippy','--offline','--locked','--manifest-path',manifest,'--all-targets','--all-features','--','-D','warnings']),('doc',['cargo','doc','--offline','--locked','--manifest-path',manifest,'--all-features','--no-deps']),('test',['cargo','test','--offline','--locked','--manifest-path',manifest,'--all-features','--','--test-threads=1'])]
receipts=[]
for name,args in checks:
 label=f'{quality_label}-probe-{name}'
 code=d.run(label,args);receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
 if code:sys.exit(code)
d.write(d.P/f'probe-quality-{leg}.json',dict(status='pass',receipts=receipts))
for kind in ['native','allocation']:
 freeze=dict(source=source,probe=probe,driver=d.desc(d.P/'driver.py'),builder=d.desc(d.P/'build_probe.py'))
 d.write(d.P/f'freeze-{leg}-{kind}.json',freeze)
 bin_name='cfb-emission-probe-alloc' if kind=='allocation' else 'cfb-emission-probe'
 args=['cargo','build','--offline','--locked','--release','--manifest-path',manifest,'--bin',bin_name]
 label=f'build-{leg}-{kind}'
 assert d.run(label,args)==0
 binary=d.TARGET/'retained'/('leg-a' if leg=='before' else 'leg-b')/kind/'cfb-emission-probe';binary.parent.mkdir(parents=True,exist_ok=False)
 shutil.copy2(d.TARGET/'release'/bin_name,binary)
 assert d.source()==source and all(d.sha(d.P/n)==h for n,h in probe.items())
 d.write(d.P/f'build-{leg}-{kind}.json',dict(binary=d.desc(binary),freeze=d.desc(d.P/f'freeze-{leg}-{kind}.json'),receipt=d.desc(d.P/'commands'/label/'receipt.json')))
