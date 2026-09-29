"""Validate the final ZIP write retry fix after rejecting the OPC optimization."""
import driver as d
packages=['soapberry-zip','litchi-opc','litchi-docx','litchi-xlsx','litchi-pptx']
selectors=[arg for package in packages for arg in ['-p',package]]
commands=[
 ('final-fmt',['cargo','fmt','--all','--','--check']),
 ('final-check',['cargo','check','--offline','--locked','--all-targets',*selectors]),
 ('final-clippy',['cargo','clippy','--offline','--locked','--all-targets',*selectors,'--','-D','warnings']),
 ('final-doc',['cargo','doc','--offline','--locked','--no-deps',*selectors]),
 ('final-tests',['cargo','test','--offline','--locked','--no-fail-fast',*selectors]),
 ('final-boundaries',['python3','-B','tools/check_crate_boundaries.py']),
]
source=d.source();d.write(d.P/'final-quality-source.json',dict(source=source))
for label,argv in commands:
 assert d.source()==source
 assert d.run(label,argv)==0,label
assert d.source()==source
d.write(d.P/'final-quality.json',dict(status='pass',source=source,commands=[d.desc(d.P/f'commands/{label}/receipt.json') for label,_ in commands]))
print('final quality PASS',flush=True)
