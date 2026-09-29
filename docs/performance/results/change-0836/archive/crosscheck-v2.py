"""Independent small owner-frame census from mmap records and ELF segments."""
import gzip,re
import driver as d

def main():
 symbols=d.read(d.P/'symbols.json');stage=symbols['stages']['fp']
 mappings=d.read(d.P/'mappings.json');elf=open(mappings['elf']['path']).read()
 loads=[]
 for line in elf.splitlines():
  f=line.split()
  if f and f[0]=='LOAD' and 'E' in f:
   loads.append((int(f[1],16),int(f[2],16),int(f[4],16)))
 assert len(loads)==1
 offset,vaddr,size=loads[0];results=[]
 for row,mrow in zip(d.read(d.P/'profiles.json')['rows'],mappings['rows']):
  assert row['label']==mrow['label']
  report=d.read(row['report']['path'])
  pids={s['child_process_id'] for e in report['filesystem_evidence'] for s in e['samples']}
  maps={}
  for line in gzip.open(mrow['decoded']['path'],'rt'):
   if 'PERF_RECORD_MMAP2' not in line or not line.rstrip().endswith(stage['binary']['path']):continue
   m=re.search(r'(\d+\.\d+): PERF_RECORD_MMAP2 (\d+)/\d+: \[0x([0-9a-f]+)\(0x([0-9a-f]+)\) @ 0x([0-9a-f]+) <([^>]+)>\]: r-xp ',line)
   assert m,line
   t,pid,start,length,off,buildid=m.groups();pid=int(pid)
   assert buildid==stage['build_id'] and int(off,16)==offset//4096*4096
   bias=int(start,16)-(vaddr//4096*4096)
   assert pid not in maps
   maps[pid]=(float(t),int(start,16),int(start,16)+int(length,16),bias)
  assert pids<=maps.keys()
  count=0;period_sum=0;named_outside=0;pid=None;period=0;t=0;qualified=False
  def finish():return int(qualified),period if qualified else 0
  for line in gzip.open(row['decoded']['path'],'rt'):
   h=re.match(r'^\S+\s+(\d+)\s+(\d+\.\d+):\s+(\d+) cycles:u:',line)
   if h:
    c,p=finish();count+=c;period_sum+=p
    pid,t,period=int(h[1]),float(h[2]),int(h[3]);qualified=False
   elif pid in pids and symbols['owner']+'+0x' in line and line.rstrip().endswith('('+stage['binary']['path']+')'):
    addr=int(line.split()[0],16);mt,start,end,bias=maps[pid]
    assert t>=mt and start<=addr<end
    adjusted=addr-bias
    hits=[r for r in stage['ranges'] if r['address']<=adjusted<r['end']]
    if hits:qualified=True
    else:named_outside+=1
  c,p=finish();count+=c;period_sum+=p
  results.append(dict(label=row['label'],owner_samples=count,owner_period=period_sum,owner_named_outside_ranges=named_outside,measured_pids=len(pids)))
 d.write(d.P/'owner-crosscheck.json',dict(status='pass',rows=results,scope='Independent mmap/ELF/static-range owner census; no CPU or wall fraction.'))
 print(results)

if __name__=='__main__':main()
