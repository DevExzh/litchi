"""Admit the actual optimized publication scope before observer capture."""
import driver as d
import re
b=d.read(d.P/'build.json')['binary'];assert d.desc(b['path'])==b
records={}
for name,argv in [('raw',['nm','-S','--defined-only',b['path']]),('demangled',['nm','-S','--defined-only','-C',b['path']]),('elf',['readelf','-n',b['path']])]:
 label='symbols-'+name;assert d.run(label,argv)==0;records[name]=d.desc(d.P/'commands'/label/'output.log')
def rows(item):
 result=[]
 for line in d.Path(item['path']).read_text().splitlines():
  m=re.fullmatch(r'([0-9a-fA-F]+)\s+([0-9a-fA-F]+)\s+([A-Za-z])\s+(.+)',line)
  if m:result.append(dict(address=int(m[1],16),size=int(m[2],16),kind=m[3],symbol=m[4]))
 return result
raw,dem=rows(records['raw']),rows(records['demangled']);owner=d.read(d.P/'plan.json')['profile_owner']
owned=[x for x in dem if x['kind'] in 'tT' and re.fullmatch(re.escape(owner)+r'(?:::h[0-9a-f]+)?',x['symbol'])]
assert len(owned)==1 and owned[0]['size']>0
x=owned[0];same=[y for y in raw if (y['address'],y['size'],y['kind'])==(x['address'],x['size'],x['kind'])];assert len(same)==1
assert d.run('symbols-owner-code',['objdump','-d','-C',f"--start-address={hex(x['address'])}",f"--stop-address={hex(x['address']+x['size'])}",b['path']])==0
p=d.P/'commands/symbols-owner-code/output.log';instructions=[]
for line in p.read_text().splitlines():
 m=re.match(r'^\s*([0-9a-f]+):\s+((?:[0-9a-f]{2}\s+)+)(.*)$',line)
 if m:instructions.append((int(m[1],16),m[3]))
assert instructions and instructions[0][0]==x['address'] and all(x['address']<=a<x['address']+x['size'] for a,_ in instructions)
assert any(re.search(r'\bcall\b',s) for _,s in instructions) and any(re.search(r'\bret\b',s) for _,s in instructions)
d.write(d.P/'symbols.json',dict(status='pass',binary=b,inventory=records,owner=dict(**x,raw_symbol=same[0]['symbol'],assembly=d.desc(p))))
print('Publication symbol admission PASS')
