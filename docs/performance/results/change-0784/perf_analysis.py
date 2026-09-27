"""Exact-owner native stack qualification; nested costs are not additive."""
import collections,gzip,hashlib,json,re,sys
from pathlib import Path
import custody as c
import native_analysis as n
OWNER='litchi_pptx::package::model::Package::opened_presentation_with_limits'
SCAN='litchi_pptx::notes::codec::scan_processed_xml'
FINGERPRINT='litchi_pptx::opened::model::package_fingerprint_with_memo'
RESOLVED='litchi_pptx::notes::resolved'
SHA='sha2::sha256::x86_sha::compress'

def compressed(original,zipped):
 path=n.artifact(zipped)
 with gzip.open(path,'rb') as f:data=f.read()
 assert len(data)==original['bytes'] and hashlib.sha256(data).hexdigest()==original['sha256']
 return data

def samples(data):
 result=[]
 for block in data.decode().strip().split('\n\n'):
  lines=block.splitlines();assert lines
  h=re.fullmatch(r'\S+\s+\d+\s+(\d+\.\d+):\s+(\d+) cycles:u:\s*',lines[0]);assert h,lines[0]
  symbols=[]
  for line in lines[1:]:
   m=re.fullmatch(r'\s*[0-9a-f]+ (.+?)(?:\+0x[0-9a-f]+)? \(.+\)',line);assert m,line
   symbols.append(m[1])
  result.append({'timestamp':h[1],'period':int(h[2]),'symbols':symbols})
 return result

def analyze():
 build=c.read(c.P/'build/build.json');n.binaries(build)
 fp=c.read(c.P/'build-fp/receipt.json');assert fp['exit_code']==0 and fp['environment']['RUSTFLAGS']=='-C force-frame-pointers=yes'
 n.artifact(fp['source']);n.artifact(fp['plan']);n.artifact(fp['log'])
 if Path(fp['binary']['path']).exists():n.artifact(fp['binary'])
 else:
  clean=c.read(c.P/'cleanup.json');assert clean['verified_before_removal'] and clean['binaries']['profile-fp']==fp['binary']
 answer={}
 for lane in ['perf','perf-fp']:
  out=c.P/lane;plan=c.read(c.P/('perf-plan.json' if lane=='perf' else 'perf-fp-plan.json'))
  assert plan['owner']==OWNER and plan['repeats']==2 and plan['samples']==100 and plan['warmup']==3
  assert plan['cpu']==12 and plan['frequency']==499 and plan['event']=='cycles:u'
  rows=c.read(out/'receipts.json');decode=c.read(out/'decode-receipts.json');compression=c.read(out/'compression.json')
  assert len(rows)==len(decode)==2
  packed={Path(r['original']['path']).name:r for r in compression}
  # Every retained compressed object is checked, including original exploratory decodes.
  cache={name:compressed(r['original'],r['compressed']) for name,r in packed.items()}
  frames=c.read(out/'frame-receipts.json') if lane=='perf-fp' else None
  lane_results=[]
  for i,row in enumerate(rows):
   binary=build['binaries']['profile'] if lane=='perf' else fp['binary']
   assert row['repeat']==i and row['exit_code']==0 and row['binary']==binary
   assert row['plan_sha256']==c.sha(c.P/('perf-plan.json' if lane=='perf' else 'perf-fp-plan.json'))
   assert row['raw']==packed[f'{i}.data']['original']
   n.artifact(row['log']);rp=n.artifact(row['report'])
   n.report(rp,'large','profile',100,3,{'binaries':{'profile':binary}})
   expected=['taskset','-c','12','perf','record','-e','cycles:u','-F','499','--call-graph',plan['call_graph'],'-o',row['raw']['path'],'--',binary['path'],'--mode','capture','--shape','large','--samples','100','--warmup','3','--output',row['report']['path']]
   assert row['command']==expected
   dec=decode[i];assert dec['repeat']==i and dec['exit_code']==0 and dec['raw']==row['raw']
   assert dec['command']==['perf','script','--ns','-i',row['raw']['path']]
   assert dec['decoded']==packed[f'{i}.decoded']['original'];n.artifact(dec['log'])
   stream=cache[f'{i}.decoded']
   if frames:
    fr=frames[i];assert fr['repeat']==i and fr['exit_code']==0 and fr['raw']==row['raw']
    assert fr['command']==['perf','script','--no-inline','--ns','-i',row['raw']['path']]
    n.artifact(fr['log']);stream=compressed(fr['frames'],fr['compressed'])
   all_samples=samples(stream);qualified=[s for s in all_samples if OWNER in s['symbols']]
   assert len({s['timestamp'] for s in all_samples})==len(all_samples)
   partitions={'notes_scan':0,'fingerprint':0,'other_capture':0};weighted=dict.fromkeys(partitions,0)
   nested={RESOLVED:0,SHA:0};leaf=collections.Counter();inside_unknown=0
   for s in qualified:
    names=s['symbols'];assert names.count(OWNER)==1
    interior=names[:names.index(OWNER)];assert not (SCAN in interior and FINGERPRINT in interior)
    category='notes_scan' if SCAN in interior else 'fingerprint' if FINGERPRINT in interior else 'other_capture'
    partitions[category]+=1;weighted[category]+=s['period']
    for name in nested:nested[name]+=int(name in interior)
    if interior:leaf[interior[0]]+=1
    if any('[unknown]' in v for v in interior):inside_unknown+=1
   assert sum(partitions.values())==len(qualified)
   period=sum(s['period'] for s in qualified)
   if lane=='perf-fp':assert len(qualified)>0
   lane_results.append({'repeat':i,'whole_process_samples':len(all_samples),'owner_qualified_samples':len(qualified),'unqualified_samples':len(all_samples)-len(qualified),'qualified_period_total':period,'exclusive_partition_samples':partitions,'exclusive_partition_period':weighted,'nested_inclusive_samples':nested,'top_sampled_leaf_symbols':leaf.most_common(15),'qualified_stacks_with_unknown_interior':inside_unknown})
  answer[lane]=lane_results
 return {'phase_fraction_claim_authorized':False,'fraction_refusal_reason':'Frozen plan forbids phase fractions with unresolved stacks; ordinary DWARF has zero exact owners and one FP qualified stack has an unresolved interior. Retain exact observed counts only.', 'lanes':answer,'qualification_owner':OWNER,'scope':'Only exact fully qualified owner stacks count. Native frame-pointer diagnostic and ordinary failed-unwind traces remain separate. Warmup and measured captures both included. Only observed sample and period counts are reported; no phase fractions authorized; nested costs not additive; no production speedup or population confidence claim.'}

def main():
 assert sys.argv[1:] in [['--write'],['--check']]
 result=analyze();text=json.dumps(result,indent=2,sort_keys=True)+'\n';path=c.P/'perf-analysis.json'
 if sys.argv[1]=='--write':path.write_text(text)
 else:assert path.read_text()==text
 print('Native sampled exact-owner replay PASS')
if __name__=='__main__':main()
