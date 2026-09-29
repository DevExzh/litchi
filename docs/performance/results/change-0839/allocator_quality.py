"""Fresh isolated allocator observer quality, with global-state tests serialized."""
import driver as d
import sys
leg=sys.argv[1];manifest=str(d.P/'allocator/Cargo.toml')
source=d.source();inputs={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'allocator').rglob('*')) if p.is_file()}
d.write(d.P/f'allocator-quality-{leg}-source.json',dict(source=source,inputs=inputs))
commands=[('fmt',['cargo','fmt','--manifest-path',manifest,'--','--check']),('check',['cargo','check','--offline','--locked','--all-targets','--all-features','--manifest-path',manifest]),('clippy',['cargo','clippy','--offline','--locked','--all-targets','--all-features','--manifest-path',manifest,'--','-D','warnings']),('doc',['cargo','doc','--offline','--locked','--all-features','--no-deps','--manifest-path',manifest]),('test',['cargo','test','--offline','--locked','--all-features','--manifest-path',manifest,'--','--test-threads=1'])]
for name,args in commands:
 assert d.run(f'{leg}-allocator-{name}',args)==0
assert source==d.source() and inputs=={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/'allocator').rglob('*')) if p.is_file()}
d.write(d.P/f'allocator-quality-{leg}.json',dict(status='pass',source=source,inputs=inputs,receipts=[d.desc(d.P/f'commands/{leg}-allocator-{n}/receipt.json') for n,_ in commands]))
print('allocator '+leg+' quality PASS')
