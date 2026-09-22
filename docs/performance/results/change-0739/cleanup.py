"""Remove only the owned build tree after exact binary and evidence checks."""
import json, pathlib, shutil
from run import P, ROOT, TARGET, read, sha, guard, verify_freeze
if __name__=='__main__':
    guard();verify_freeze()
    assert read('analysis.json')['status']=='passed'
    assert read('audit.json')['status']=='passed'
    assert TARGET==ROOT.parent/'litchi-target-0739' and TARGET.is_dir() and not TARGET.is_symlink()
    binaries=[]
    for b in read('build.json'):
        exe=pathlib.Path(b['binary'])
        assert exe.parent==TARGET/'release' and sha(exe)==b['binary_sha256']
        binaries.append({'path':str(exe),'sha256':sha(exe),'bytes':exe.stat().st_size})
    shutil.rmtree(TARGET)
    assert not TARGET.exists();guard()
    (P/'cleanup.json').write_text(json.dumps({'removed':str(TARGET),'binaries':binaries},indent=2)+'\n')
    print('PASS owned target cleanup after exact two-binary identity checks')
