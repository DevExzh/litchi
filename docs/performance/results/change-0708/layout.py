#!/usr/bin/env python3
"""Measure exact candidate enum slot size with the captured host compiler."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
source=ROOT/'crates/litchi-xlsx/src/cell_values/validation.rs'
text=source.read_text();start=text.index('enum ElementName {');end=text.index('\n}',start)+2
snippet='#[allow(dead_code)]\n'+text[start:end]+"\nfn main() { println!(\"{} {}\", std::mem::size_of::<ElementName>(), std::mem::size_of::<Box<[u8]>>()); }\n"
probe=P/'layout-probe.rs';assert not probe.exists();probe.write_text(snippet)
binary=ROOT.parent/'litchi-0708-bin/layout-probe'
command=['rustc','--edition=2024',str(probe),'-o',str(binary)]
result=subprocess.run(command,cwd=ROOT,capture_output=True,text=True)
(P/'layout-build.stdout').write_text(result.stdout);(P/'layout-build.stderr').write_text(result.stderr);assert result.returncode==0
output=subprocess.check_output([str(binary)],text=True);slot,old=map(int,output.split())
sha=lambda q:hashlib.sha256(q.read_bytes()).hexdigest()
(P/'layout.json').write_text(json.dumps(dict(command=command,exit_code=result.returncode,output=output,candidate_slot_bytes=slot,baseline_box_slot_bytes=old,max_xml_depth=256,logical_extra_at_max_depth=(slot-old)*256,scope='Logical slot size only, not allocation capacity, live peak or RSS',source_sha256=sha(source),snippet_sha256=sha(probe),binary=dict(path=str(binary),sha256=sha(binary),bytes=binary.stat().st_size)),indent=2)+'\n')
print(output,end='')
