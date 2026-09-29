"""Candidate integration gates and qualification before admission."""
import driver as d
assert d.source()==d.read(d.P/'source-after.json')['source']
commands=[('after-integration-quality',['python3','-B',str(d.P/'quality.py')]),('after-symbols',['python3','-B',str(d.P/'symbols.py'),'after']),('after-qualification-capture',['python3','-B',str(d.P/'capture.py'),'qualification','after']),('after-qualification-check',['python3','-B',str(d.P/'qualify.py'),'after']),('both-qualification-check',['python3','-B',str(d.P/'qualify.py')]),('after-observer-capture',['python3','-B',str(d.P/'observer_preflight.py'),'after']),('after-observer-check',['python3','-B',str(d.P/'observer_check.py'),'after'])]
for label,argv in commands:assert d.run(label,argv)==0
print('Candidate stage PASS')
