"""Complete baseline gates and qualification before candidate application."""
import driver as d
assert d.source()==d.read(d.P/'source-before.json')['source']
commands=[('before-owner-test',['cargo','test','--offline','--locked','-p','soapberry-zip','-p','litchi-opc','--all-features']),('before-symbols',['python3','-B',str(d.P/'symbols.py'),'before']),('before-qualification-capture',['python3','-B',str(d.P/'capture.py'),'qualification','before']),('before-qualification-check',['python3','-B',str(d.P/'qualify.py'),'before']),('before-observer-capture',['python3','-B',str(d.P/'observer_preflight.py'),'before']),('before-observer-check',['python3','-B',str(d.P/'observer_check.py'),'before'])]
for label,argv in commands:assert d.run(label,argv)==0
previous=d.P.parent/'change-0844'
assert d.read(previous/'source-after.json')['source']==d.source()
assert d.read(d.P/'commands/baseline-replay/receipt.json')['exit_code']==0
d.write(d.P/'baseline-gate-reuse.json',dict(status='pass',production=d.desc(d.P/'source-before.json'),prior_quality=d.desc(previous/'quality-after.json'),prior_closure=d.desc(previous/'closure.json'),prior_seal=d.desc(previous/'seal.json'),fresh_replay=d.desc(d.P/'commands/baseline-replay/receipt.json'),fresh_owner_tests=d.desc(d.P/'commands/before-owner-test/receipt.json')))
print('Baseline stage PASS')
