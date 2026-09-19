#!/usr/bin/env python3
"""Temporary tracing identifies remaining walks; never use its timings as evidence."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];out=P/'route-trace';out.mkdir(exist_ok=True)
paths=[ROOT/'crates/litchi-cfb/src/shared.rs',ROOT/'crates/litchi-xls/src/workbook/source.rs'];original={p:p.read_bytes() for p in paths}
base=json.loads((P/'baseline.json').read_text())['baseline_head']
def sha(b):return hashlib.sha256(b).hexdigest()
for p,b in original.items():assert b==subprocess.check_output(['git','show',base+':'+str(p.relative_to(ROOT))],cwd=ROOT)
commands=[];target=Path('/home/zhuhe/code/litchi-target-0689');binary=target/'release/xls0684-repeat'
try:
 p=paths[0];s=p.read_text();old='    let (mut sector, walked) = resume;\n';assert s.count(old)==2
 s=s.replace(old,'    eprintln!("TRACE chain table={table_name} from={} to={ordinal} links={}", resume.1, ordinal.saturating_sub(resume.1));\n'+old,1);p.write_text(s)
 p=paths[1];s=p.read_text();old='    let mut strings = SharedStringResolver::new(owner, &refs);\n    let mut chain = index.chain_checkpoint';assert s.count(old)==1
 s=s.replace(old,'    eprintln!("TRACE replay row={row} column={column}");\n'+old)
 old='    let first_offset = sst_source_offset(&segments[first_segment], location.start)?;';assert s.count(old)==1
 s=s.replace(old,old+'\n    eprintln!("TRACE sst index={string_index} first_offset={first_offset}");')
 start=s.index('fn replay_indexed_cell(');end=s.index('\n/// Reads and decodes one indexed occurrence.',start);part=s[start:end];old='        let mut cursor = owner';assert part.count(old)==1
 part=part.replace(old,'        eprintln!("TRACE worksheet frame_offset={} sheet_start={}", slot.stream_offset, sheet.start);\n'+old);s=s[:start]+part+s[end:];p.write_text(s)
 (out/'instrumentation.patch').write_bytes(subprocess.check_output(['git','diff','--',*[str(p.relative_to(ROOT)) for p in paths]],cwd=ROOT))
 instrumented={str(p.relative_to(ROOT)):sha(p.read_bytes()) for p in paths}
 cmd=['cargo','build','--manifest-path','docs/performance/results/change-0684/repeat-probe/Cargo.toml','--release','--locked','--offline'];start=time.monotonic()
 with (out/'build.log').open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(target),CARGO_BUILD_JOBS='2'),stdout=f,stderr=subprocess.STDOUT)
 commands.append(dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start));assert r.returncode==0
 binary_hash=sha(binary.read_bytes())
 for c in json.loads((P/'cases.json').read_text()):
  if c['budget']!=2097152:continue
  cmd=[str(binary),'owned',c['path'],str(c['sheet']),str(c['row']),str(c['column']),'2']
  with (out/(c['case']+'.tsv')).open('w') as o,(out/(c['case']+'.trace')).open('w') as e:r=subprocess.run(cmd,cwd=ROOT,stdout=o,stderr=e)
  expected=1 if c['case'].startswith('formula-refusal') else 0
  commands.append(dict(command=cmd,exit_code=r.returncode,expected_exit_code=expected));assert r.returncode==expected
 (out/'manifest.json').write_text(json.dumps(dict(baseline_head=base,instrumented_source_sha256=instrumented,binary_sha256=binary_hash,commands=commands,scope='Diagnostic tracing only; stderr instrumentation invalidates performance timings.',raw_sha256={f.name:sha(f.read_bytes()) for f in out.iterdir() if f.is_file() and f.name!='manifest.json'}),indent=2)+'\n')
finally:
 for p,b in original.items():p.write_bytes(b)
 assert all(p.read_bytes()==b for p,b in original.items())
print('Route traces captured; original sources restored.',flush=True)
