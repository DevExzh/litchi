#!/usr/bin/env python3
"""Separate fixed ABBA plain-tail investigation; primary samples remain untouched."""
import hashlib,importlib.util,json,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def module(name):
 s=importlib.util.spec_from_file_location(name,ROOT/(name+'.py'));m=importlib.util.module_from_spec(s);s.loader.exec_module(m);return m
d=module('derive');oracle=module('verify-report')
def sha(raw):return hashlib.sha256(raw).hexdigest()
def artifact(r):
 p=Path(r['path']);d.require(not p.is_absolute() and '..' not in p.parts,'confirmation safe artifact');raw=d.raw(r['path']);d.require(len(raw)==r['bytes'] and sha(raw)==r['sha256'],'confirmation artifact');return raw
def distribution(samples):
 v=[s['timings']['api_sum_ns']/1e6 for s in samples]
 return {'samples':len(v),'api_p50_ms':statistics.median(v),'api_p95_ms':d.quantile(v,.95),'api_p99_ms':d.quantile(v,.99),'api_max_ms':max(v),'phase_p50_ms':{p:statistics.median(s['timings'][p+'_ns']/1e6 for s in samples) for p in ['open','plan','publication']}}
def derive():
 rows=[];samples={b:[] for b in ['baseline','candidate']};blocks=[]
 protocol=d.load('confirmation-protocol.json')
 for i,lane in enumerate(protocol['order']):
  report=d.load(f'confirmation/{i}/report.json');v=report['samples_raw'];samples[lane['build']]+=v
  rows.append({'lane':i,**lane,**distribution(v)})
 for k in range(2):
  pair={b:distribution([s for i in range(k*4,k*4+4) if protocol['order'][i]['build']==b for s in d.load(f'confirmation/{i}/report.json')['samples_raw']]) for b in samples}
  blocks.append({'block':k+1,**pair,'percent_change':{m:d.delta(pair['baseline'][m],pair['candidate'][m]) for m in ['api_p50_ms','api_p95_ms','api_p99_ms']}})
 pooled={b:distribution(v) for b,v in samples.items()}
 return {'change':453,'scope':'separate two-ABBA plain/bytes tail investigation; per-block 60 and overall 120 samples per build combine process distributions, not population p99 estimates; primary reports unchanged','samples':240,'rows':rows,'blocks':blocks,'pooled':pooled,'pooled_percent_change':{m:d.delta(pooled['baseline'][m],pooled['candidate'][m]) for m in ['api_p50_ms','api_p95_ms','api_p99_ms']}}
def render(v):
 lines=['# Separate plain-copy tail investigation','','Eight fresh processes, two fixed ABBA blocks, three warmups and thirty samples.','Primary matrix remains unchanged. Combined quantiles are descriptive; these','sample counts cannot establish a stable population p99.','','| Lane | Build | p50 ms | p95 ms | p99 ms | Maximum ms |','|---|---|---:|---:|---:|---:|']
 for r in v['rows']:lines.append(f"| {r['lane']} | {r['build']} | {r['api_p50_ms']:.6f} | {r['api_p95_ms']:.6f} | {r['api_p99_ms']:.6f} | {r['api_max_ms']:.6f} |")
 lines+=['','| ABBA block | p50 change | p95 change | p99 change |','|---|---:|---:|---:|']
 for r in v['blocks']:lines.append('| '+str(r['block'])+' | '+' | '.join(f"{r['percent_change'][k]:+.3f}%" for k in ['api_p50_ms','api_p95_ms','api_p99_ms'])+' |')
 lines+=['','Pooled descriptive changes: '+json.dumps(v['pooled_percent_change'],sort_keys=True)+'.']
 return '\n'.join(lines)+'\n'
def check(final,builds):
 protocol=d.load('confirmation-protocol.json');order=[dict(provider='bytes',corpus='plain',build=b,instrumentation='normal',repeat='C'+str(i//4+1)) for i,b in enumerate(['baseline','candidate','candidate','baseline']*2)]
 d.require(protocol['order']==order and protocol['samples']==30 and protocol['warmups']==3 and protocol['cpu']==2 and protocol['reports']==8 and protocol['status']=='frozen','fixed confirmation protocol')
 for n,h in protocol['bound_files'].items():d.require(sha(d.raw(n))==h,'confirmation frozen dependency')
 paths=list((ROOT/'confirmation').glob('*/receipt.json'));d.require(len(paths)==8,'confirmation inventory')
 for p in paths:
  i=int(p.parent.name);r=json.loads(p.read_text());lane=order[i];build=builds[lane['build']];binary=build['binaries']['normal']
  d.require(r['lane']==lane and r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0 and not r['pilot'] and r['profile'] is None,'confirmation pass')
  d.require(r['source_unchanged'] and r['source_before']==r['source_after']==builds['candidate']['source_manifest'],'confirmation source')
  d.require(r['binary']==binary and r['build_sha256']==sha(d.raw(lane['build']+'-build.json')) and r['protocol_sha256']==sha(d.raw('confirmation-protocol.json')) and r['capture_sha256']==sha(d.raw('confirmation-capture.py')) and r['oracle_sha256']==sha(d.raw('verify-report.py')),'confirmation custody')
  for a in r['artifacts'].values():artifact(a)
  report=json.loads(artifact(r['artifacts']['report']));oracle.check_report(report)
  d.require(report['binary_sha256']==binary['sha256'] and report['binary_bytes']==binary['bytes'] and report['current_exe']==binary['path'] and report['source_revision']==build['revision'],'confirmation executable')
  d.require(report['provider']=='bytes' and report['corpus']=='plain' and report['samples']==30 and report['warmup']==3 and report['instrumentation']=='none','confirmation workload')
  for k,v in protocol['corpora']['plain'].items():d.require(report[k]==v,'confirmation exact corpus/output')
  argv=r['argv'];d.require(argv==['taskset','-c','2','/usr/bin/time','-v','-o',argv[6],binary['path'],'provider-lifecycle','--corpus','plain','--provider','bytes','--samples','30','--warmup','3','--source-revision',build['revision'],'--output',argv[-1]],'confirmation exact command')
 v=derive();d.require(v==d.load('confirmation-summary.json') and render(v)==d.raw('confirmation-summary.md').decode(),'confirmation derivation')
 return v
if __name__=='__main__':
 v=derive();(ROOT/'confirmation-summary.json').write_text(json.dumps(v,indent=2)+'\n');(ROOT/'confirmation-summary.md').write_text(render(v));print(json.dumps(v['pooled_percent_change']))
