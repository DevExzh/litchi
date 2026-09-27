"""Offline analysis of paired probe samples, exact output parity and RSS."""
from pathlib import Path
import hashlib,json,statistics
P=Path(__file__).resolve().parent
M=P/'measure-1'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def analyze():
 runs=json.loads((M/'runs.json').read_text());assert len(runs)==16
 assert json.loads((M/'complete.json').read_text())['source_unchanged']
 bins=json.loads((M/'binaries.json').read_text());assert bins['before']['lock_sha256']==bins['after']['lock_sha256']
 rows=[];previous=0
 for r in runs:
  assert r['exit']==0 and r['started']>=previous;previous=r['ended']
  for key in ['stdout','stderr','rss']:assert sha(M/r[key])==r[key+'_sha256']
  data=json.loads((M/r['stdout']).read_text());assert data['case']==r['case'] and data['cells']==20000 and data['samples']==9 and data['warmup']==2
  samples=data['samples_ns'];assert len(samples)==9 and all(x>0 for x in samples)
  assert r['command'][r['command'].index('-c')+1]=='12'
  binary=r['command'][r['command'].index('-c')+2];assert binary==bins[r['leg']]['path']
  ordered=sorted(samples)
  rows.append({'case':r['case'],'leg':r['leg'],'file':r['stdout'],'p50_ms':statistics.median(samples)/1e6,'p95_ms':ordered[-1]/1e6,'p99_ms':ordered[-1]/1e6,'rss_kib':int((M/r['rss']).read_text()),'output_bytes':data['output_bytes'],'output_sha256':data['output_sha256']})
 cases={}
 for case in ['formula','numeric']:
  selected=[r for r in rows if r['case']==case];assert len({(r['output_bytes'],r['output_sha256']) for r in selected})==1
  legs={}
  for leg in ['before','after']:
   s=[r for r in selected if r['leg']==leg];assert len(s)==4
   p50s=[r['p50_ms'] for r in s];rss=[r['rss_kib'] for r in s]
   legs[leg]={'median_process_p50_ms':statistics.median(p50s),'p50_range_ms':[min(p50s),max(p50s)],'p50_spread_pct':100*(max(p50s)/min(p50s)-1),'median_process_peak_rss_kib':statistics.median(rss),'rss_range_kib':[min(rss),max(rss)],'rss_spread_pct':100*(max(rss)/min(rss)-1)}
  latency=legs['after']['median_process_p50_ms']/legs['before']['median_process_p50_ms'];memory=legs['after']['median_process_peak_rss_kib']/legs['before']['median_process_peak_rss_kib']
  cases[case]={'legs':legs,'after_before_latency_ratio':latency,'after_before_peak_rss_ratio':memory,'latency_regression_over_5pct':latency>1.05,'rss_regression_over_5pct':memory>1.05,'process_spread_over_5pct':any(x['p50_spread_pct']>5 for x in legs.values()),'rss_process_spread_over_5pct':any(x['rss_spread_pct']>5 for x in legs.values()),'output_bytes':selected[0]['output_bytes'],'output_sha256':selected[0]['output_sha256']}
 return {'scope':'descriptive four-process-per-leg per-case release probe; nine samples/process; p95/p99 are nine-sample maxima, not stable tail estimates','cases':cases,'processes':rows}
if __name__=='__main__':print(json.dumps(analyze(),indent=2))
