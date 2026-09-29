"""Fresh standalone diagnostic quality; production source-identical gates retained."""
import driver as d
manifest=str(d.P/'probe/Cargo.toml')
commands=[
 ('probe-v2-fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),
 ('probe-v2-check',['cargo','check','--offline','--locked','--all-targets','--manifest-path',manifest]),
 ('probe-v2-clippy',['cargo','clippy','--offline','--locked','--all-targets','--manifest-path',manifest,'--','-D','warnings']),
 ('probe-v2-doc',['cargo','doc','--offline','--locked','--no-deps','--manifest-path',manifest]),
]
source=d.source();probe={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/'quality-v2-source.json',dict(source=source,probe=probe))
for label,argv in commands:
 assert d.source()==source
 assert d.run(label,argv)==0,label
assert d.source()==source
assert probe=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'probe').rglob('*')) if p.is_file()}
d.write(d.P/'quality.json',dict(status='pass',source=source,probe=probe,commands=[d.desc(d.P/f'commands/{label}/receipt.json') for label,_ in commands]))
print('probe quality PASS',flush=True)
