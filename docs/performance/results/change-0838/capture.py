"""Three serial, untimed fresh-process diagnostic captures."""
import driver as d
build=d.read(d.P/'build-diagnostic.json');binary=build['binary']
assert d.desc(binary['path'])==binary
freeze=d.read(d.P/'freeze-diagnostic.json')
assert d.source()==freeze['source']
d.write(d.P/'admission.json',{'files':[d.desc(d.P/n) for n in ['plan.json','quality.json','quality-reuse.json','build-diagnostic.json','probe/src/main.rs','probe/Cargo.toml','probe/Cargo.lock','capture.py','audit.py','fixtures/fresh-source.zip']]})
for i in range(3):
 dest=d.P/'runs'/f'run-{i:02}'
 assert not dest.exists()
 assert d.run(f'capture-{i:02}',['taskset','-c','12',binary['path'],str(d.P/'fixtures/fresh-source.zip'),str(dest)])==0
assert d.source()==freeze['source']
d.write(d.P/'capture.json',dict(status='pass',processes=3,operations=18,timed_operations=0,artifacts=[d.desc(p) for p in sorted((d.P/'runs').rglob('*')) if p.is_file()]))
