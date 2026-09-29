"""Admit exact existing owner symbols before any profiling capture."""
import re,sys
import driver as d

def rows(path):
 result=[]
 for line in path.read_text().splitlines():
  m=re.fullmatch(r'([0-9a-fA-F]+)\s+([0-9a-fA-F]+)\s+([A-Za-z])\s+(.+)',line)
  if m:result.append(dict(address=int(m[1],16),size=int(m[2],16),kind=m[3],symbol=m[4]))
 return result

def collect(stage):
 binary=d.read(d.P/f'build-{stage}.json')['binary'];assert d.desc(binary['path'])==binary
 records={}
 for name,argv in [('raw',['nm','-S','--defined-only',binary['path']]),('demangled',['nm','-S','--defined-only','-C',binary['path']]),('elf',['readelf','-n',binary['path']])]:
  label=f'symbols-{stage}-{name}';r=d.run(stage,label,argv);assert r['exit_code']==0
  records[name]=dict(output=d.desc(d.P/f'commands/{label}/output.log'),receipt=d.desc(d.P/f'commands/{label}/receipt.json'))
 d.write(d.P/f'symbols-{stage}.json',dict(stage=stage,binary=binary,records=records))
 print('symbol inventory',stage,'PASS')

def admit():
 inventories={stage:d.read(d.P/f'symbols-{stage}.json') for stage in ['baseline','fp']}
 parsed={stage:{name:rows(d.P/f'commands/symbols-{stage}-{name}/output.log') for name in ['raw','demangled']} for stage in inventories}
 candidates=d.read(d.P/'plan.json')['perf_contract']['owner_candidates']
 def matches(symbol,candidate):return symbol==candidate or re.fullmatch(re.escape(candidate)+r'::h[0-9a-f]+',symbol) is not None
 selected=None;selected_rows=None
 for candidate in candidates:
  matching={s:[r for r in parsed[s]['demangled'] if r['kind'] in 'Tt' and matches(r['symbol'],candidate)] for s in inventories}
  if all(matching.values()):selected=candidate;selected_rows=matching;break
 assert selected is not None,'no exact owner exists in both builds; profiling not admitted'
 stages={}
 for stage,demangled in selected_rows.items():
  binary=inventories[stage]['binary'];ranges=[];end=0
  for i,row in enumerate(sorted(demangled,key=lambda r:r['address'])):
   assert row['size']>0 and row['address']>=end;end=row['address']+row['size']
   raw=[r for r in parsed[stage]['raw'] if (r['address'],r['size'],r['kind'])==(row['address'],row['size'],row['kind'])]
   assert len(raw)==1,'ambiguous raw owner symbol'
   label=f'symbols-{stage}-owner-{i:02}'
   argv=['objdump','-d','-C',f"--start-address={hex(row['address'])}",f"--stop-address={hex(end)}",binary['path']]
   r=d.run(stage,label,argv);assert r['exit_code']==0
   path=d.P/f'commands/{label}/output.log';text=path.read_text()
   instructions=[]
   for line in text.splitlines():
    m=re.match(r'^\s*([0-9a-f]+):\s+((?:[0-9a-f]{2}\s+)+)(.*)$',line)
    if m:instructions.append((int(m[1],16),m[3]))
   assert instructions and instructions[0][0]==row['address']
   assert all(row['address']<=a<end for a,_ in instructions)
   prefix='\n'.join(t for _,t in instructions[:10])
   if stage=='fp':assert re.search(r'push\s+%rbp',prefix) and re.search(r'mov\s+%rsp,%rbp',prefix),'frame pointer prologue missing'
   assert any(re.search(r'\bcall\b',t) for _,t in instructions),'owner has no nested call'
   ranges.append(dict(**row,raw_symbol=raw[0]['symbol'],end=end,assembly=d.desc(path),receipt=d.desc(d.P/f'commands/{label}/receipt.json'),frame_pointer_prologue=stage=='fp'))
  elf=d.P/f'commands/symbols-{stage}-elf/output.log';ids=re.findall(r'Build ID:\s*([0-9a-f]+)',elf.read_text());assert len(ids)==1
  stages[stage]=dict(binary=binary,inventory=d.desc(d.P/f'symbols-{stage}.json'),build_id=ids[0],ranges=ranges)
 scope='whole source-backed save function' if selected==candidates[0] else 'source-backed single-Part overlay publication only; excludes package open and final atomic synchronization'
 d.write(d.P/'symbols.json',dict(status='pass',owner=selected,scope=scope,stages=stages,plan=d.desc(d.P/'plan.json')))
 print('owner admitted:',selected,'scope:',scope)
if __name__=='__main__':
 if sys.argv[1]=='collect':collect(sys.argv[2])
 else:assert sys.argv[1:] == ['admit'];admit()
