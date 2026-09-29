"""Existing affected-owner and format-dependent quality gates."""
import driver as d

packages=['soapberry-zip','litchi-opc','litchi-docx','litchi-xlsx','litchi-pptx']
selectors=[arg for package in packages for arg in ['-p',package]]
commands=[
 ('candidate-fmt',['cargo','fmt','--all','--','--check']),
 ('probe-fmt',['cargo','fmt','--manifest-path',str(d.P/'probe/Cargo.toml'),'--','--check']),
 ('candidate-check',['cargo','check','--offline','--locked','--all-targets',*selectors]),
 ('candidate-clippy',['cargo','clippy','--offline','--locked','--all-targets',*selectors,'--','-D','warnings']),
 ('candidate-doc',['cargo','doc','--offline','--locked','--no-deps',*selectors]),
 ('candidate-tests',['cargo','test','--offline','--locked',*selectors]),
 ('candidate-boundaries',['python3','-B','tools/check_crate_boundaries.py']),
]
for label,argv in commands:
 assert d.run(label,argv)==0,label
d.write(d.P/'quality.json',dict(status='pass',source=d.source(),commands=[d.desc(d.P/f'commands/{label}/receipt.json') for label,_ in commands]))
print('candidate quality PASS',flush=True)
