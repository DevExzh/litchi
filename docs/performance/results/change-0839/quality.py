"""Root-only scoped quality gates, executed serially for each source leg."""
import driver as d
import sys
leg=sys.argv[1];manifest=str(d.P/'probe/Cargo.toml')
commands=[
 ('fmt',['cargo','fmt','--all','--','--check']),
 ('check',['cargo','check','--offline','--locked','-p','litchi-opc','--all-features','--all-targets']),
 ('clippy',['cargo','clippy','--offline','--locked','-p','litchi-opc','--all-features','--lib','--','-D','warnings']),
 ('doc',['cargo','doc','--offline','--locked','-p','litchi-opc','--all-features','--no-deps']),
 ('tests',['cargo','test','--offline','--locked','-p','litchi-opc','--all-features']),
 ('boundaries',['python3','-B','tools/check_crate_boundaries.py']),
 ('probe-fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),
 ('probe-check',['cargo','check','--offline','--locked','--all-targets','--all-features','--manifest-path',manifest]),
 ('probe-clippy',['cargo','clippy','--offline','--locked','--all-targets','--all-features','--manifest-path',manifest,'--','-D','warnings']),
 ('probe-doc',['cargo','doc','--offline','--locked','--all-features','--no-deps','--manifest-path',manifest]),
 ('probe-tests',['cargo','test','--offline','--locked','--all-features','--manifest-path',manifest]),
]
source=d.source();probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/f'quality-{leg}-source.json',dict(source=source,probe=probe))
for label,argv in commands:
 assert d.source()==source
 assert d.run(leg+'-'+label,argv)==0,label
assert d.source()==source
assert probe=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/f'quality-{leg}.json',dict(status='pass',source=source,probe=probe,commands=[d.desc(d.P/f'commands/{leg}-{label}/receipt.json') for label,_ in commands]))
print(leg+' quality PASS',flush=True)
