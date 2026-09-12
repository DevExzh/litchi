"""Build both allocator variants serially under their exact frozen source."""
import subprocess
from run import HERE, REPO, TARGET, build, capture, check_source, plan_data

check_source('candidate')
codec = REPO/'crates/litchi-ooxml-common/src/mce/codec.rs'
test = REPO/'crates/litchi-ooxml-common/tests/mce_namespace_search.rs'
backup = TARGET/'candidate-source-backup'
backup.mkdir(exist_ok=False)
(backup/'codec.rs').write_bytes(codec.read_bytes())
(backup/'test.rs').write_bytes(test.read_bytes())
try:
    codec.write_bytes(subprocess.check_output([
        'git', 'show', plan_data()['revision']+':crates/litchi-ooxml-common/src/mce/codec.rs'], cwd=REPO))
    test.unlink()
    check_source('baseline')
    build('baseline', 'alloc')
finally:
    codec.write_bytes((backup/'codec.rs').read_bytes())
    test.write_bytes((backup/'test.rs').read_bytes())
check_source('candidate')
capture('baseline', 'alloc')
build('candidate', 'alloc')
capture('candidate', 'alloc')
