"""Remove only this packet's owned target after checking both executable hashes."""
import shutil, time
import custody as c
assert c.TARGET == __import__('pathlib').Path(c.read(c.P/'origin.json')['target'])
assert c.TARGET.name == 'litchi-target-0786' and c.TARGET.is_dir() and not c.TARGET.is_symlink()
assert not (c.P/'cleanup.json').exists()
build=c.read(c.P/'build.json');checked={}
for kind,expected in build['binaries'].items():
    path=__import__('pathlib').Path(expected['path'])
    assert path.parent==c.TARGET and not path.is_symlink()
    actual=c.artifact(path);assert actual==expected
    checked[kind]=actual
files=[p for p in c.TARGET.rglob('*') if p.is_file()]
record={'verified':True,'executables_verified_before_removal':True,'binaries':checked,'target':str(c.TARGET),'file_count':len(files),'file_bytes':sum(p.stat().st_size for p in files),'started':time.time()}
shutil.rmtree(c.TARGET)
record.update({'ended':time.time(),'target_absent_after_removal':not c.TARGET.exists()})
c.write(c.P/'cleanup.json',record)
print(record['file_bytes'],'owned target file bytes removed')
