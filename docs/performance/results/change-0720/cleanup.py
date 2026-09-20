#!/usr/bin/env python3
"""Remove only the finished 0720 build target and record its binary identity."""
import json,shutil
from pathlib import Path
from custody import P,ROOT,TARGET,census,sha
assert census()==json.loads((P/'baseline-source.json').read_text())
quality=json.loads((P/'quality.json').read_text())
assert len(quality)==8 and all(r['exit_code']==0 for r in quality)
scratch=ROOT.parent/'litchi-scratch-0720'
assert not scratch.exists(), 'restore production instrumentation first'
assert TARGET==Path('/home/zhuhe/code/litchi-target-0720') and not TARGET.is_symlink()
binary=TARGET/'debug/docx-active-offset-oracle-0720'
identity={'path':str(binary),'bytes':binary.stat().st_size,'sha256':sha(binary)}
assert identity['sha256']==json.loads((P/'trace-build-receipt.json').read_text())['binary_sha256']
assert not (P/'cleanup.json').exists()
shutil.rmtree(TARGET)
record={'removed_paths':[str(TARGET),str(scratch)],'trace_binary':identity,'baseline_binary_disposition':'replaced by successful trace build after its baseline capture','production_source_restored':True}
assert all(not Path(n).exists() for n in record['removed_paths'])
(P/'cleanup.json').write_text(json.dumps(record,indent=2)+'\n')
print('owned target removed; production source and scratch restoration verified')
