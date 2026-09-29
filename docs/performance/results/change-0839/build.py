"""Retain separately identified native, source-observer and allocator binaries."""
import driver as d
import shutil,sys
leg=sys.argv[1]
if (d.P/f'build-{leg}.json').exists():
 assert d.source()==d.read(d.P/f'freeze-{leg}.json')['source']
 assert d.desc(d.read(d.P/f'build-{leg}.json')['binary']['path'])==d.read(d.P/f'build-{leg}.json')['binary']
else:d.build(leg)
source=d.source()
for kind,manifest,features,name in [('observer','probe',['--features','source-metrics'],'cached-part-memory'),('allocator','allocator',[],'cached-part-memory-alloc')]:
 if len(sys.argv)>2 and kind not in sys.argv[2:]:continue
 argv=['cargo','build','--offline','--locked','--release','--manifest-path',str(d.P/manifest/'Cargo.toml'),*features]
 d.write(d.P/f'freeze-{leg}-{kind}.json',dict(source=source,inputs={str(p.relative_to(d.P)):d.sha(p) for p in sorted((d.P/manifest).rglob('*')) if p.is_file()}))
 assert d.run(f'build-{leg}-{kind}',argv)==0
 assert d.source()==source
 binary=d.TARGET/'retained'/leg/kind/name;binary.parent.mkdir(parents=True,exist_ok=False);shutil.copy2(d.TARGET/'release'/name,binary)
 d.write(d.P/f'build-{leg}-{kind}.json',dict(binary=d.desc(binary),freeze=d.desc(d.P/f'freeze-{leg}-{kind}.json'),receipt=d.desc(d.P/f'commands/build-{leg}-{kind}/receipt.json')))
