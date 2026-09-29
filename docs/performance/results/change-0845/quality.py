"""Affected OPC/OOXML and workspace library integration gates."""
import driver as d
source=d.source();assert source==d.read(d.P/'source-after.json')['source'];d.check()
packages=['soapberry-zip','litchi-opc','litchi-docx','litchi-xlsx','litchi-pptx']
owners=[v for name in packages for v in ['-p',name]]
commands=[('fmt',['cargo','fmt','--all','--check']),('workspace-check',['cargo','check','--offline','--locked','--workspace','--all-features','--lib','--tests']),('owner-check',['cargo','check','--offline','--locked',*owners,'--all-features','--all-targets']),('workspace-clippy',['cargo','clippy','--offline','--locked','--workspace','--all-features','--lib','--no-deps','--','-D','warnings']),('owner-clippy',['cargo','clippy','--offline','--locked',*owners,'--all-features','--all-targets','--no-deps','--','-D','warnings']),('owner-doc',['cargo','doc','--offline','--locked',*owners,'--all-features','--no-deps']),('owner-test',['cargo','test','--offline','--locked',*owners,'--all-features']),('boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
receipts=[]
for name,args in commands:
 label='after-'+name;assert d.run(label,args)==0;receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
assert d.source()==source
d.write(d.P/'quality-after.json',dict(status='pass',receipts=receipts,source=d.desc(d.P/'source-after.json'),excluded='iWork examples/full-workspace rustdoc retain baseline-confirmed exclusions; workspace lib/tests check and lib Clippy included.'))
