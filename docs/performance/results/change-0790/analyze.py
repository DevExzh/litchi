"""Offline launcher qualification; preserve counter scope and every raw child."""
import gzip,re,sys
from collections import Counter
import common as c


def kv(raw,key):
 m=re.search(r'^'+re.escape(key)+r':\s+(\d+)(?:\s+kB)?$',raw,re.M);assert m,key
 return int(m[1])


def binary(a):
 if c.resolve(a).exists():c.verify(a)
 else:
  clean=c.read(c.P/'cleanup.json');assert clean['target_removed'] and clean['target']==str(c.TARGET) and not c.TARGET.exists()
  assert clean['removed_binaries'].count(a)==1


def check():
 plan=c.read(c.P/'plan.json');build=c.read(c.P/'build.json');frozen=c.read(c.verify(build['inputs']))
 for n,h in frozen.items():assert c.sha(c.P/n)==h,n
 inherited=c.read(c.P/'inherited.json')
 for a in inherited.values():c.verify(a)
 assert c.sha(c.P/'probe.c')==inherited['probe.c']['sha256']
 for key in ['binary','launcher','sanitized_launcher']:binary(build[key])
 rows=[];ack=Counter();checks=[];failures=[]
 for lane in ['matrix','traces']:
  complete=c.read(c.P/lane/'complete.json');receipts=c.read(c.verify(complete['receipts']));assert len(receipts)==complete['children']==(96 if lane=='matrix' else 4)
  expected=[]
  if lane=='matrix':
   for repeat,order in enumerate(plan['orders']):
    for case in plan['cases']:
     for affinity in plan['affinities']:
      for label in order:
       launcher,observer=label.split('-');expected.append((case,affinity,launcher,observer,repeat))
  else:
   expected=[({k:item[k] for k in ['mib','workers','touch']},item['affinity'],'direct','full',0) for item in plan['traces']]
  for index,(artifact,key) in enumerate(zip(receipts,expected)):
   r=c.read(c.verify(artifact));case,affinity,launcher,observer,repeat=key
   assert r['schema']=='litchi.fork-launcher-child.0790.v1' and r['index']==index
   assert (r['case'],r['affinity'],r['launcher'],r['observer'],r['repeat'])==key
   assert r['trace']==(lane=='traces') and r['error'] is None and r['wait4']['exit_code']==0
   assert not c.verify(r['stderr']).read_bytes() and r['binary']==build['binary']
   cpus=plan['affinities'][affinity]
   assert r['command'][:3]==['/usr/bin/taskset','-c',','.join(map(str,cpus))]
   assert r['command'][-7:]==[build['binary']['path'],'--mib',str(case['mib']),'--workers',str(case['workers']),'--touch',case['touch']]
   ps=__import__('json').loads(gzip.decompress(c.verify(r['points']).read_bytes()));lines=c.verify(r['transcript']).read_text().splitlines();assert len(ps)==6 and len(lines)==12
   child=ps[0]['identity']['pid'];parent=ps[0]['identity']['ppid'];stamp=ps[0]['identity']['starttime']
   usage=c.read(c.verify(r['fork_usage'])) if launcher=='direct' else None
   if usage:
    assert usage['child_pid']==child and usage['exit_code']==usage['wait_status']==0
    exit_usage=usage['rusage']
    if ps[0]['before']['maxrss_kib']>=usage['launcher_start_maxrss_kib']:
     failures.append({'lane':lane,'index':index,'criterion':'direct startup lower than launcher pre-fork high water','child_startup_kib':ps[0]['before']['maxrss_kib'],'launcher_startup_kib':usage['launcher_start_maxrss_kib']})
    assert ps[0]['before']['maxrss_kib']<17708
   else:
    vals=list(map(int,c.verify(r['gnu_time']).read_text().split()));assert len(vals)==3
    exit_usage=dict(zip(['maxrss_kib','minor_faults','major_faults'],vals))
   snapshots=[];self_values=[]
   pages=case['mib']*1024*1024//4096;checksum=sum(i%251+1 for i in range(pages))
   for i,(p,phase) in enumerate(zip(ps,plan['phases'])):
    assert p['phase']==phase
    ident=p['identity'];assert ident=={'pid':child,'ppid':parent,'starttime':stamp,'exe':build['binary']['path'],'task_ids':[child]}
    pi=p['parent_identity'];assert pi['pid']==parent and pi['exe']==(build['launcher']['path'] if launcher=='direct' else '/usr/bin/time')
    assert (pi['ppid'] if lane=='traces' else parent)==r['wrapper_pid']
    stat=p['raw']['stat'];tail=stat[stat.rfind(')')+2:].split();assert int(stat.split()[0])==child and int(tail[1])==parent and int(tail[17])==1 and int(tail[19])==stamp and int(tail[36]) in cpus
    before=lines[2*i].split('\t');after=lines[2*i+1].split('\t')
    assert len(before)==8 and before[:3]==['RSS0789',phase,str(child)]
    assert len(after)==6 and after[:3]==['ACK0789',phase,str(child)]
    assert p['before']==dict(zip(['maxrss_kib','minor_faults','major_faults','mapped_bytes','checksum'],map(int,before[3:])))
    assert p['after']==dict(zip(['maxrss_kib','minor_faults','major_faults'],map(int,after[3:])))
    assert p['before']['mapped_bytes']==(case['mib']*1024*1024 if i in [1,2,3] else 0)
    assert p['before']['checksum']==(checksum if i>=2 else 0)
    for k in ['maxrss_kib','minor_faults','major_faults']:assert p['after'][k]>=p['before'][k]>=0
    self_values += [p['before']['maxrss_kib'],p['after']['maxrss_kib']]
    if lane=='matrix':ack[(launcher,observer,p['after']['maxrss_kib']-p['before']['maxrss_kib'])]+=1
    if observer=='full':
     assert set(p['raw'])=={'stat','smaps','smaps_rollup','maps','status'}
     roll=p['raw']['smaps_rollup'];status=p['raw']['status'];rss=kv(roll,'Rss')
     assert rss==sum(map(int,re.findall(r'^Rss:\s+(\d+)\s+kB$',p['raw']['smaps'],re.M)))==kv(status,'VmRSS')
     assert kv(status,'Pid')==child and kv(status,'PPid')==parent and kv(status,'Threads')==1
     assert re.search(r'^Cpus_allowed_list:\s+(.+)$',status,re.M)[1].strip()==('0' if affinity=='single' else '0-31')
     snapshots.append({'phase':phase,'rss_kib':rss,'anonymous_kib':kv(roll,'Anonymous'),'vmhwm_kib':kv(status,'VmHWM')})
    else:assert set(p['raw'])=={'stat'}
   row={k:r[k] for k in ['index','case','affinity','launcher','observer','repeat']}
   row.update(lane=lane,startup_self_kib=self_values[0],max_self_kib=max(self_values),exit_usage=exit_usage,launcher_start_maxrss_kib=usage['launcher_start_maxrss_kib'] if usage else None,outer_wait4_scope='strace wrapper' if lane=='traces' else 'launcher wrapper',outer_wait4=r['wait4']['rusage'],snapshots=snapshots,exit_minus_max_self_kib=exit_usage['maxrss_kib']-max(self_values))
   if snapshots:
    growth=snapshots[2]['anonymous_kib']-snapshots[1]['anonymous_kib'];row.update(anon_growth_kib=growth,anon_extra_kib=growth-case['mib']*1024,observed_minus_exit_kib=max(x['rss_kib'] for x in snapshots)-exit_usage['maxrss_kib'])
    if case['workers']==0:assert row['anon_extra_kib']==0
   if lane=='traces':
    trace=c.verify(r['strace']).read_text().splitlines()
    def counters(line):return [int(re.search(n+r'=(\d+)',line)[1]) for n in ['ru_maxrss','ru_minflt','ru_majflt']]
    calls=[line for line in trace if re.match(str(child)+r'\s+getrusage\(RUSAGE_SELF,',line)];assert len(calls)==12
    for line,u in zip(calls,[p[k] for p in ps for k in ['before','after']]):assert counters(line)==[u[k] for k in ['maxrss_kib','minor_faults','major_faults']]
    waits=[line for line in trace if re.match(str(parent)+r'\s+.*wait4',line) and re.search(r' = '+str(child)+r'$',line)];assert len(waits)==1 and 'WEXITSTATUS(s) == 0' in waits[0]
    assert counters(waits[0])==[exit_usage[k] for k in ['maxrss_kib','minor_faults','major_faults']]
    checks.append({'index':index,'self_samples':12,'wait4':exit_usage})
   rows.append(row)
 qual=c.read(c.P/'qualification/checks.json');assert len(qual)==7
 for r in qual:
  c.verify(r['stdout']);c.verify(r['stderr']);assert r['exit_code']==r['expected']
  if 'usage' in r:assert c.read(c.verify(r['usage']))['exit_code']==r['expected']
 for r in build['commands']:assert r['exit_code']==0;assert not c.verify(r['log']).read_bytes()
 matrix=[r for r in rows if r['lane']=='matrix'];groups=[]
 for launcher in ['direct','time']:
  for affinity in plan['affinities']:
   selected=[r for r in matrix if r['launcher']==launcher and r['affinity']==affinity]
   groups.append({'launcher':launcher,'affinity':affinity,'children':len(selected),'startup_self_range_kib':[min(r['startup_self_kib'] for r in selected),max(r['startup_self_kib'] for r in selected)],'exit_equal_max_self':sum(r['exit_minus_max_self_kib']==0 for r in selected),'full_observed_minus_exit_range_kib':[min(r['observed_minus_exit_kib'] for r in selected if r['observer']=='full'),max(r['observed_minus_exit_kib'] for r in selected if r['observer']=='full')]})
 return {'schema':'litchi.fork-launcher-analysis.0790.v1','acceptance':'fail' if failures else 'pass','acceptance_failures':failures,'scope':'diagnostic-only; no adoption, confidence intervals or transient peak claim','children':len(rows),'matrix_children':96,'trace_children':4,'full_snapshots':312,'identity_snapshots':288,'qualification_checks':7,'rows':rows,'groups':groups,'trace_checks':checks,'matrix_ack_deltas':[{'launcher':k[0],'observer':k[1],'delta_kib':k[2],'count':v} for k,v in sorted(ack.items())]}

if __name__=='__main__':
 result=check();path=c.P/'analysis.json'
 if '--check' in sys.argv:assert c.read(path)==result;print('0790 replay PASS: 100 children, 600 checkpoints, seven qualifications')
 else:assert not path.exists();c.write(path,result);print('0790 analysis written')
