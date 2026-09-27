"""Remove only the owned target after verifying all six retained executable identities."""
import shutil,time
from pathlib import Path
import custody as c
assert c.TARGET==Path(c.read(c.P/'origin.json')['target'])
assert c.TARGET.name=='litchi-target-0787' and c.TARGET.is_dir() and not c.TARGET.is_symlink()
assert not (c.P/'cleanup.json').exists()
expected={}
for leg in ('before','after'):
    for kind,record in c.read(c.P/f'build-{leg}/build.json')['binaries'].items():
        expected[f'{leg}-{kind}']=record
for name,folder in [('baseline-profile','profile-build'),('baseline-profile-compat','profile-build-compat')]:
    expected[name]=c.read(c.P/f'{folder}/receipt.json')['binary']
assert len(expected)==6
checked={}
for key,record in expected.items():
    path=Path(record['path']);assert path.parent==c.TARGET and not path.is_symlink()
    assert c.artifact(path)==record
    checked[key]=record
files=[p for p in c.TARGET.rglob('*') if p.is_file()]
record={'verified':True,'executables_verified_before_removal':True,'binaries':checked,'target':str(c.TARGET),'file_count':len(files),'file_bytes':sum(p.stat().st_size for p in files),'started':time.time()}
shutil.rmtree(c.TARGET)
record.update({'ended':time.time(),'target_absent_after_removal':not c.TARGET.exists()})
c.write(c.P/'cleanup.json',record)
print(record['file_bytes'],'owned target file bytes removed')
