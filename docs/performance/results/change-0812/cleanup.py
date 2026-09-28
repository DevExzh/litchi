"""Remove only the target recreated for the exact historical binary rebuild."""
import shutil
import time
import inputs as c

TARGET = c.ROOT.parent / 'litchi-target-0811'
assert TARGET == __import__('pathlib').Path('/home/zhuhe/code/litchi-target-0811')
assert TARGET.is_dir() and not TARGET.is_symlink()
assert not (c.P / 'cleanup.json').exists()
assert (c.P / 'instruction-analysis.json').is_file()
assert (c.P / 'offset-audit.json').is_file()
assert (c.P / 'fresh-offset-audit.json').is_file()
witness = c.read(c.P / 'fresh/complete.json')['binary']
assert c.artifact(TARGET / 'fp') == witness
custody = c.verify()
paths = [path for path in TARGET.rglob('*') if path.is_file()]
count, size = len(paths), sum(path.stat().st_size for path in paths)
started = time.time()
shutil.rmtree(TARGET)
assert not TARGET.exists()
assert c.verify() == custody
receipt = {'schema': 'litchi.performance.0812.cleanup.v1', 'target': str(TARGET),
           'removed_files': count, 'removed_logical_bytes': size, 'binary': witness,
           'inputs': custody, 'started': started, 'ended': time.time(), 'target_removed': True}
(c.P / 'cleanup.json').write_text(__import__('json').dumps(receipt, indent=2, sort_keys=True) + '\n')
print('0812 recreated target removed:', count, 'files;', size, 'logical bytes')
