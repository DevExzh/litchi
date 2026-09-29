"""Serial shared-owner and OOXML integration gates."""
import driver as d
import sys
leg=sys.argv[1];prefix=leg+'-final';d.check();source=d.source()
d.write(d.P/f'quality-{prefix}-source.json',dict(source=source))
commands=[('fmt',['cargo','fmt','--all','--check']),('workspace-check',['cargo','check','--offline','--locked','--workspace','--all-features','--lib','--tests']),('owner-check',['cargo','check','--offline','--locked','-p','soapberry-zip','-p','litchi-opc','-p','litchi-docx','-p','litchi-xlsx','-p','litchi-pptx','--all-features','--all-targets']),('workspace-clippy',['cargo','clippy','--offline','--locked','--workspace','--all-features','--lib','--no-deps','--','-D','warnings']),('workspace-doc',['cargo','doc','--offline','--locked','--workspace','--all-features','--no-deps']),('owner-test',['cargo','test','--offline','--locked','-p','soapberry-zip','-p','litchi-opc','-p','litchi-docx','-p','litchi-xlsx','-p','litchi-pptx','--all-features']),('boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
receipts=[]
for name,args in commands:
 label=f'{prefix}-{name}';code=d.run(label,args);receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
 if code:sys.exit(code)
assert d.source()==source
d.write(d.P/f'quality-{leg}.json',dict(status='pass',receipts=receipts,source=source))
