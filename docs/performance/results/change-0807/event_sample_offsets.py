"""Post-capture exploratory leaf-IP census; sampled IPs are not cost fractions."""
import collections,gzip,hashlib,re,sys
import custody as c
symbol='quick_xml::reader::ns_reader::NsReader<R>::process_event'
owner='litchi_pptx::package::model::Package::opened_presentation_with_limits'
assembly=c.read(c.P/'event-assembly.json')
assert assembly['binary']==c.read(c.P/'build-fp/receipt.json')['binary']
for key in ['symbols','assembly']:
 assert c.artifact(assembly[key]['path'])==assembly[key]
base=assembly['address'];size=assembly['size']
assert assembly['nm_exit_code']==assembly['objdump_exit_code']==0
assert assembly['nm_command']==['nm','-S','--defined-only',assembly['binary']['path']]
assert assembly['objdump_command']==['objdump','-d','--demangle','--no-show-raw-insn',f'--start-address={base}',f'--stop-address={base+size}',assembly['binary']['path']]
assert assembly['started']<=assembly['ended']
for tool,receipt in c.read(c.P/'event-tools.json').items():
 assert tool in ['nm','objdump'] and receipt['command']==[tool,'--version'] and receipt['exit_code']==0
entries=[line.split() for line in (c.P/'event-symbols.txt').read_text().splitlines() if line.endswith(' '+assembly['symbol'])]
assert len(entries)==1 and int(entries[0][0],16)==base and int(entries[0][1],16)==size
instructions={int(m[1],16)-base:m[2] for line in (c.P/'event-assembly.txt').read_text().splitlines() if (m:=re.fullmatch(r'\s*([0-9a-f]+):\s+(.+)',line))}
rows=[]
for receipt,audit in zip(c.read(c.P/'perf-fp/frame-receipts.json'),c.read(c.P/'root-frame-counts.json')['rows'],strict=True):
 assert receipt['repeat']==audit['repeat']
 z=receipt['compressed'];assert c.artifact(z['path'])==z
 raw=gzip.decompress(__import__('pathlib').Path(z['path']).read_bytes())
 assert len(raw)==receipt['frames']['bytes'] and hashlib.sha256(raw).hexdigest()==receipt['frames']['sha256']
 counts=collections.Counter()
 for block in raw.decode().strip().split('\n\n'):
  frames=[]
  for line in block.splitlines()[1:]:
   m=re.fullmatch(r'\s*[0-9a-f]+ (.+?)(?:\+0x([0-9a-f]+))? \((.+)\)',line)
   assert m,line
   frames.append((m[1],int(m[2] or '0',16),m[3]))
  names=[f[0] for f in frames]
  if names.count(owner)!=1 or not frames or frames[0][0]!=symbol:continue
  name,offset,dso=frames[0]
  assert dso==assembly['binary']['path'] and 0<=offset<size and offset in instructions
  counts[offset]+=1
 assert sum(counts.values())==dict(audit['top_leaf'])[symbol]
 rows.append({'repeat':receipt['repeat'],'leaf_samples':sum(counts.values()),'offsets':[{'offset_hex':hex(k),'samples':v,'instruction':instructions[k]} for k,v in sorted(counts.items())]})
value={'scope':'Exploratory post-capture sampled leaf IP offsets for one exact frame-pointer binary, including warmups. Sampling skid, attribution and code-generation limits apply; no instruction cost fractions or causal speedup claim.','binary':assembly['binary'],'symbol':symbol,'rows':rows}
p=c.P/'event-sample-offsets.json'
if '--check' in sys.argv:assert c.read(p)==value
else:assert not p.exists();c.write(p,value)
print('Exploratory event leaf-IP census PASS',flush=True)
