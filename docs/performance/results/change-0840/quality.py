"""Run fresh affected-owner gates serially and retain every terminal receipt."""
import driver as d
import sys
leg=sys.argv[1];label_prefix=sys.argv[2] if len(sys.argv)>2 else leg
d.check();source=d.source()
d.write(d.P/f'quality-{label_prefix}-source.json',dict(source=source))
commands=[('fmt',['cargo','fmt','--all','--check']),('check',['cargo','check','--offline','--locked','-p','litchi-cfb','-p','litchi-doc','--all-features','--all-targets']),('clippy',['cargo','clippy','--offline','--locked','-p','litchi-cfb','-p','litchi-doc','--all-features','--lib','--','-D','warnings']),('doc',['cargo','doc','--offline','--locked','-p','litchi-cfb','-p','litchi-doc','--all-features','--no-deps']),('test',['cargo','test','--offline','--locked','-p','litchi-cfb','-p','litchi-doc','--all-features']),('boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
receipts=[]
for name,args in commands:
 label=f'{label_prefix}-{name}'
 code=d.run(label,args)
 receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
 if code:sys.exit(code)
assert d.source()==source
d.write(d.P/f'quality-{leg}.json',dict(status='pass',receipts=receipts,source=source))
