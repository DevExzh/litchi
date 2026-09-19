#!/usr/bin/env python3
"""Inspect actual checked-link and cursor-construction code in each binary."""
import hashlib,json,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent;phase=sys.argv[1];suffix='before' if phase=='baseline' else 'after'
binary=Path('/home/zhuhe/code/litchi-target-0690-'+suffix)/'release/xls0684-repeat';symbols=subprocess.check_output(['nm','-S','-C','--defined-only',str(binary)],text=True);rows=[]
for label,symbol in [('next-chain-sector','litchi_cfb::shared::next_chain_sector'),('cursor','litchi_cfb::shared::SharedOleFile::stream_cursor_at_hinted'),('chain-walk','litchi_cfb::shared::cursor_chain_sector'),('marker-error','litchi_cfb::shared::invalid_chain_marker'),('index-error','litchi_cfb::shared::invalid_chain_index'),('query-cell','litchi_xls::workbook::source::query_cell'),('directory-name','litchi_cfb::directory_name::directory_name_data'),('shared-lookup','litchi_cfb::shared::SharedOleFile::find_child_by_name'),('ascii-lookup','litchi_cfb::directory_name::ascii_lookup_key')]:
 matches=[l.split(maxsplit=3) for l in symbols.splitlines() if l.split(maxsplit=3)[-1]==symbol]
 if not matches:
  assert label in ['next-chain-sector','shared-lookup','ascii-lookup'];rows.append(dict(label=label,symbol=symbol,standalone_symbol=False));continue
 assert len(matches)==1;address,size,kind,name=matches[0];cmd=['objdump','-d','-C','--start-address=0x'+address,'--stop-address='+hex(int(address,16)+int(size,16)),str(binary)]
 f=P/(phase+'-'+label+'-assembly.txt');f.write_text(subprocess.check_output(cmd,text=True));rows.append(dict(label=label,symbol=symbol,standalone_symbol=True,command=cmd,code_bytes=int(size,16),assembly_file=f.name,assembly_sha256=hashlib.sha256(f.read_bytes()).hexdigest()))
size_output=subprocess.check_output(['size','-A',str(binary)],text=True)
(P/(phase+'-binary-sections.txt')).write_text(size_output)
sections={line.split()[0]:int(line.split()[1]) for line in size_output.splitlines() if line.startswith('.')}
(P/(phase+'-assembly-manifest.json')).write_text(json.dumps(dict(binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),symbols=rows,section_bytes=sections,section_file=phase+'-binary-sections.txt',section_sha256=hashlib.sha256(size_output.encode()).hexdigest()),indent=2)+'\n')
