"""Independent exact-input/result and raw paired-native numerical audit."""
import json,random,re,statistics,sys
from pathlib import Path
p=Path(__file__).resolve().parent
read=lambda n:json.loads(Path(n).read_text())
MASK=(1<<64)-1;STEP=0x9e3779b97f4a7c15;SEED=0x6a09e667f3bcc909;VALUE=0xbb67ae8584caa73b
rot=lambda x,n:((x<<n)|(x>>(64-n)))&MASK

def fnv(data):
 h=0xcbf29ce484222325
 for v in data:h=((h^v)*0x100000001b3)&MASK
 return h

def expected_case(row):
 case=row['id'];n=int(case.rsplit('-',1)[1]);prefix='e'+''.join(f' n{i}="{i}"' for i in range(n));error=None
 if case.startswith('distinct-'):source=prefix
 elif case.startswith('duplicate-'):
  if case.startswith('duplicate-valid-'):tail='n0="again"'
  elif case.startswith('duplicate-long-quoted-'):tail='n0="'+'x'*4096+'"'
  elif case.startswith('duplicate-long-unterminated-'):tail='n0="'+'x'*4096
  else:assert case.startswith('duplicate-unquoted-');tail='n0='+'x'*4096
  source=prefix+' '+tail;error={'kind':'Duplicated','position':len(prefix)+1,'first_position':2}
 else:
  tail={'syntax-flag':'flag','syntax-unique-tail':'tail=x','syntax-equals-value':'="1"'}[case.rsplit('-after-',1)[0]]
  source=prefix+' '+tail
  if tail=='tail=x':error={'kind':'UnquotedValue','position':len(prefix)+6}
  else:error={'kind':'ExpectedEq','position':len(source)}
 raw=row['source'];observed=raw['value'].encode() if raw['encoding']=='utf8' else bytes.fromhex(raw['value'])
 assert observed==source.encode() and raw['bytes']==len(observed)
 trace=row['expected_baseline'];assert trace['accepted']==n and trace['first_error']==error and trace['repeated_none_calls']==2
 items=[]
 for i in range(n):
  k=f'n{i}'.encode();v=str(i).encode();items.append({'kind':'Attribute','key_bytes':len(k),'key_hash':fnv(k),'value_bytes':len(v),'value_hash':fnv(v)})
 if error:items.append({'kind':'Error','error':error})
 assert trace['items']==items
 one=VALUE
 for i in range(n):one=rot(one,7)^(rot(VALUE,5)^((len(f'n{i}')*STEP)&MASK)^((len(str(i))*0xd6e8feb86659fd93)&MASK))
 marker=0
 if error:
  kind={'ExpectedEq':1,'ExpectedValue':2,'UnquotedValue':3,'ExpectedQuote':4,'Duplicated':5}[error['kind']]
  marker=((kind*STEP)&MASK)^rot(error['position'],17)^rot(error.get('first_position',error.get('quote',0)),31)
 return n,one,marker

cases={r['id']:r for r in read(p/'cases.json')};assert len(cases)==39
expected={case:expected_case(row) for case,row in cases.items()}
values={};counts=0
for lane in ['native','profiles']:
 for receipt in read(p/lane/'receipts.json'):
  r=read(receipt['report']['path']);case,mode,leg=r['case'],r['mode'],r['leg'];n,one,marker=expected[case];iterations=r['iterations']
  total=(SEED+sum((i*STEP)&MASK if mode=='construct' else one^((i*STEP)&MASK) for i in range(1,iterations+1)))&MASK
  answer={'checksum':total,'accepted':0 if mode=='construct' else n*iterations,'error_marker':0 if mode=='construct' else (marker*iterations)&MASK}
  assert r['expected_result']==answer
  for sample in r['samples']:assert all(sample[k]==v for k,v in answer.items());counts+=1
  if lane=='native':
   samples=sorted(s['elapsed_ns'] for s in r['samples']);assert len(samples)==30
   values[(case,mode,receipt['block'],leg)]=samples[14]
assert counts==28392 and len(values)==936
rows=[]
for case in cases:
 for mode in ['construct','consume']:
  ratios=[values[(case,mode,b,'after')]/values[(case,mode,b,'before')] for b in range(6)]
  rng=random.Random(802080);draws=sorted(statistics.median(ratios[rng.randrange(6)] for _ in range(6)) for _ in range(10000))
  median=statistics.median(ratios)
  rows.append({'case':case,'mode':mode,'ratio':median,'ci_low':draws[250],'ci_high':draws[9749],'regression':median>1.05 and draws[250]>1.0})
result={'reports':1248,'samples':counts,'rows':rows,'production_adoption':False}
out=p/'root-native-audit.json'
if '--check' in sys.argv:assert read(out)==result
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print('Independent input/checksum and native numerical audit PASS')
