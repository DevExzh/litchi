"""Admit exact wrapper and retain unmodified optimized code for both legs."""
import re,sys
import driver as d
leg=sys.argv[1];b=d.read(d.P/f'build-{leg}.json')['binary'];assert d.desc(b['path'])==b
records={}
for name,argv in [('raw',['nm','-S','--defined-only',b['path']]),('demangled',['nm','-S','--defined-only','-C',b['path']]),('elf',['readelf','-n',b['path']])]:
 label=f'symbols-{leg}-{name}';assert d.run(label,argv)==0
 records[name]=d.desc(d.P/'commands'/label/'output.log')
def rows(path):
 result=[]
 for line in d.Path(path).read_text().splitlines():
  m=re.fullmatch(r'([0-9a-fA-F]+)\s+([0-9a-fA-F]+)\s+([A-Za-z])\s+(.+)',line)
  if m:result.append(dict(address=int(m[1],16),size=int(m[2],16),kind=m[3],symbol=m[4]))
 return result
raw=rows(records['raw']['path']);dem=rows(records['demangled']['path'])
owner=d.read(d.P/'plan.json')['profile_owner']
match=lambda s,n:s==n or re.fullmatch(re.escape(n)+r'::h[0-9a-f]+',s) is not None
owned=[r for r in dem if r['kind'] in 'tT' and match(r['symbol'],owner)]
assert len(owned)==1 and owned[0]['size']>0
# These mechanism functions may inline; absence is reported, never treated as zero work.
queries=[owner,'litchi_ppt::writer::core::codec::convert_shape_to_escher_with_sound_mapping','litchi_ppt::writer::escher::codec::text::append_client_textbox_with_interactions','litchi_ppt::writer::escher::codec::shapes::append_shape_container_children','litchi_ppt::writer::escher::codec::validation::validate_user_shape']
selected=[]
for query in queries:
 for row in [r for r in dem if r['kind'] in 'tT' and match(r['symbol'],query)]:
  rraw=[r for r in raw if (r['address'],r['size'],r['kind'])==(row['address'],row['size'],row['kind'])];assert len(rraw)==1
  label=f'symbols-{leg}-code-{len(selected):02}'
  argv=['objdump','-d','-C',f"--start-address={hex(row['address'])}",f"--stop-address={hex(row['address']+row['size'])}",b['path']]
  assert d.run(label,argv)==0
  p=d.P/'commands'/label/'output.log';instructions=[]
  for line in p.read_text().splitlines():
   m=re.match(r'^\s*([0-9a-f]+):\s+((?:[0-9a-f]{2}\s+)+)(.*)$',line)
   if m:instructions.append((int(m[1],16),m[3]))
  assert instructions and instructions[0][0]==row['address'] and all(row['address']<=a<row['address']+row['size'] for a,_ in instructions)
  if query==owner:
   assert any(re.search(r'\bcall\b',t) for _,t in instructions),'wrapper must retain a nested call'
   assert any(re.search(r'\bret\b',t) for _,t in instructions),'wrapper must return rather than only tail-call'
  selected.append(dict(**row,query=query,raw_symbol=rraw[0]['symbol'],assembly=d.desc(p)))
d.write(d.P/f'symbols-{leg}.json',dict(status='pass',binary=b,inventory=records,owner=owner,selected=selected,absent_queries=[q for q in queries if not any(r['query']==q for r in selected)]))
print('symbols',leg,'PASS')
