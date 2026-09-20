#!/usr/bin/env python3
"""Remove only this packet's owned roots after terminal capture and validation."""
import datetime,json,os,shutil
from pathlib import Path
from custody import P,ROOT,TARGET,BIN,census,sha
assert not (P/'cleanup.json').exists()
assert census()==json.loads((P/'source-final.json').read_text())
rows=json.loads((P/'capture.json').read_text());assert len(rows)==20 and all(r['exit_code']==0 for r in rows)
for lane,count in [('docx',5),('harness',1),('evidence',6)]:
 checks=json.loads((P/f'quality-{lane}.json').read_text());assert len(checks)==count and all(r['exit_code']==0 for r in checks)
for n in ['analysis.json','read-controls-analysis.json','trace-analysis.json','negative-checks.json']:assert (P/n).is_file()
roots=[TARGET,BIN,ROOT.parent/'litchi-0722-fs',ROOT.parent/'litchi-scratch-0722-trace']
active=[]
for path in Path('/proc').glob('[0-9]*/cmdline'):
 try:args=path.read_bytes().split(b'\0');exe=os.fsdecode(args[0])
 except (OSError,IndexError):continue
 if exe.startswith(str(BIN)+'/') or exe.startswith(str(TARGET)+'/') or Path(exe).name in ['cargo','rustc']:active.append({'pid':path.parent.name,'exe':exe})
assert not active,active
binaries=[]
for lane in ['baseline','candidate']:
 for kind in ['native','alloc','oracle','trace']:
  path=BIN/(lane+'-'+kind);assert path.is_file() and not path.is_symlink()
  binaries.append({'path':str(path),'sha256':sha(path),'bytes':path.stat().st_size})
record={'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'owned_paths':list(map(str,roots)),'binaries':binaries,'live_cargo_or_benchmark_processes':active,'owned_paths_absent':False}
(P/'cleanup.json').write_text(json.dumps(record,indent=2)+'\n')
for path in roots:
 assert path.parent==ROOT.parent and not path.is_symlink()
 if path.exists():shutil.rmtree(path)
record['owned_paths_absent']=all(not path.exists() for path in roots);assert record['owned_paths_absent']
(P/'cleanup.json').write_text(json.dumps(record,indent=2)+'\n')
for path in P.rglob('__pycache__'):shutil.rmtree(path)
print('Removed owned build, binary, filesystem and trace roots; retained eight exact binary witnesses')
