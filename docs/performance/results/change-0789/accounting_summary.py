"""Independent raw-counter summary and exact verbose-strace cross-check."""
import gzip,json,re,sys
from collections import Counter
import common as c


def field(raw,name):
 match=re.search(r'^'+re.escape(name)+r':\s+(\d+)\s+kB$',raw,re.M)
 assert match,name
 return int(match[1])


def summarize():
 rows=[];ack=Counter();traces=[]
 for lane in ['matrix','traces','trace-detail']:
  complete=c.read(c.P/lane/'complete.json');receipts=c.read(c.verify(complete['receipts']));assert len(receipts)==complete['children']
  for index,artifact in enumerate(receipts):
   r=c.read(c.verify(artifact));assert r['index']==index and r['error'] is None and r['wait4']['exit_code']==0
   ps=json.loads(gzip.decompress(c.verify(r['points']).read_bytes()));assert len(ps)==6
   assert not c.verify(r['stderr']).read_bytes()
   time_values=list(map(int,c.verify(r['gnu_time']).read_text().split())) if r['launcher']=='time' else None
   exit_rss=time_values[0] if time_values else r['wait4']['rusage']['maxrss_kib']
   peak_self=max(p[k]['maxrss_kib'] for p in ps for k in ['before','after'])
   for p in ps:
    if lane=='matrix':ack[(r['observer'],p['after']['maxrss_kib']-p['before']['maxrss_kib'])]+=1
   row={k:r[k] for k in ['index','case','affinity','launcher','observer','repeat']}
   row.update(lane=lane,startup_self_kib=ps[0]['before']['maxrss_kib'],max_self_kib=peak_self,exit_rss_kib=exit_rss,exit_minus_max_self_kib=exit_rss-peak_self)
   if r['observer']=='full':
    rss=[field(p['raw']['smaps_rollup'],'Rss') for p in ps];anon=[field(p['raw']['smaps_rollup'],'Anonymous') for p in ps]
    row.update(touched_rss_kib=rss[2],max_observed_rss_kib=max(rss),observed_minus_exit_kib=max(rss)-exit_rss,anon_growth_kib=anon[2]-anon[1],anon_growth_beyond_payload_kib=anon[2]-anon[1]-r['case']['mib']*1024)
    for p,rr in zip(ps,rss):
     assert sum(map(int,re.findall(r'^Rss:\s+(\d+)\s+kB$',p['raw']['smaps'],re.M)))==rr
     assert field(p['raw']['status'],'VmRSS')==rr
    if r['case']['workers']==0:assert row['anon_growth_beyond_payload_kib']==0
   if lane=='trace-detail':
    lines=c.verify(r['strace']).read_text().splitlines();pid=ps[0]['identity']['pid'];parent=ps[0]['identity']['ppid']
    samples=[line for line in lines if re.match(str(pid)+r'\s+getrusage\(RUSAGE_SELF,',line)]
    assert len(samples)==12
    def counters(line):
     return [int(re.search(name+r'=(\d+)',line)[1]) for name in ['ru_maxrss','ru_minflt','ru_majflt']]
    for line,p in zip(samples,[p[k] for p in ps for k in ['before','after']]):
     assert counters(line)==[p[k] for k in ['maxrss_kib','minor_faults','major_faults']]
    waits=[line for line in lines if re.match(str(parent)+r'\s+.*wait4',line) and re.search(r' = '+str(pid)+r'$',line)]
    assert len(waits)==1 and 'WEXITSTATUS(s) == 0' in waits[0]
    assert counters(waits[0])==time_values
    traces.append({'index':index,'case':r['case'],'affinity':r['affinity'],'pid':pid,'gnu_time':time_values,'traced_wait4':counters(waits[0]),'self_samples_matched':12,'trace':r['strace']})
   rows.append(row)
 matrix=[r for r in rows if r['lane']=='matrix'];assert len(matrix)==272
 groups=[]
 for launcher in ['direct','time']:
  for affinity in ['single','all']:
   for workers,touch in [(0,'main'),(4,'main'),(4,'workers'),(32,'main'),(32,'workers')]:
    selected=[r for r in matrix if r['launcher']==launcher and r['affinity']==affinity and r['observer']=='full' and r['case']['workers']==workers and r['case']['touch']==touch]
    def bounds(key):return [min(r[key] for r in selected),max(r[key] for r in selected)]
    groups.append({'launcher':launcher,'affinity':affinity,'workers':workers,'touch':touch,'children':len(selected),'observed_minus_exit_kib':bounds('observed_minus_exit_kib'),'anon_growth_beyond_payload_kib':bounds('anon_growth_beyond_payload_kib')})
 return {'schema':'litchi.rss-accounting-independent-summary.0789.v1','children':len(rows),'matrix_children':len(matrix),'primary_trace_children':4,'supplement_trace_children':4,'full_snapshots':864,'identity_snapshots':816,'rows':rows,'groups':groups,'ack_deltas':[{'observer':k[0],'after_minus_before_kib':k[1],'checkpoints':v} for k,v in sorted(ack.items())],'exit_equal_max_self':{launcher:sum(r['exit_minus_max_self_kib']==0 for r in matrix if r['launcher']==launcher) for launcher in ['direct','time']},'verbose_trace_checks':traces,'scope':'No latency, adoption, transient peak or confidence-interval claim. Launcher pre-exec high-water differs; full and identity observations remain separate.'}

if __name__=='__main__':
 result=summarize();path=c.P/'accounting-summary.json'
 if '--check' in sys.argv:assert c.read(path)==result;print('Independent accounting summary PASS')
 else:assert not path.exists();c.write(path,result);print('Independent accounting summary written')
