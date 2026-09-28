"""Verify the three owned executables, then remove only this batch target."""
import shutil
import custody as c
assert not (c.P/'cleanup.json').exists()
assert c.TARGET==c.ROOT.parent/'litchi-target-0807'
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
assert c.source()['files']==c.read(c.P/'build/source.json')['files']
for name in ['native/complete.json','profiles/complete.json','perf-fp/complete.json',
             'perf-fp/frame-receipts.json','native-analysis.json','profile-analysis.json',
             'perf-analysis.json','root-scan-costs.json','root-frame-counts.json','quality.json']:
 assert (c.P/name).is_file(),name
build=c.read(c.P/'build/build.json')
binaries=[*build['binaries'].values(),c.read(c.P/'build-fp/receipt.json')['binary']]
assert {item['path'] for item in binaries}=={str(c.TARGET/n) for n in ['control','profile','profile-fp']}
for item in binaries:
 path=__import__('pathlib').Path(item['path'])
 assert not path.is_symlink() and c.artifact(path)==item
size=sum(path.stat().st_size for path in c.TARGET.rglob('*') if path.is_file())
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
c.write(c.P/'cleanup.json',{'target':str(c.TARGET),'target_removed':True,'removed_target_bytes':size,'removed_binaries':binaries})
print('0807 owned target and three verified binaries removed',flush=True)
