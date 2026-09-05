import pathlib,subprocess,json,os,datetime,hashlib
repo=pathlib.Path('/home/zhuhe/code/litchi');root=repo/'docs/performance/results/change-0416';out=root/'profile';out.mkdir(exist_ok=True)
env=os.environ|{'DEBUGINFOD_URLS':''};runs=[]
def run(argv,stdout=None):
 start=datetime.datetime.now(datetime.timezone.utc).isoformat()
 with open(stdout or os.devnull,'w') as log:
  p=subprocess.run(argv,stdout=log,stderr=subprocess.PIPE,text=True,env=env,cwd=out)
 runs.append({'argv':argv,'stdout':str(stdout),'exit_code':p.returncode,'stderr':p.stderr,'started_utc':start})
 (out/'capture.json').write_text(json.dumps({'scope':'separate whole-process diagnostics; no latency or allocation improvement claim','runs':runs},indent=2)+'\n')
 p.check_returncode()
fixture=str(root/'corpus/many-small-zip32-signed.zip')
for role in ('control','candidate'):
 binary='/tmp/litchi-goal-0416-'+role+'-probe';base=[binary,'index','indexed',fixture]
 data='/tmp/litchi-goal-0416-'+role+'-perf.data'
 run(['taskset','-c','2','perf','stat','-x,','-o',str(out/(role+'-stat.csv')),'-e','cycles,instructions,branches,branch-misses,page-faults','--',*base,'20000','100'],out/(role+'-stat-probe.json'))
 run(['taskset','-c','2','perf','record','-F','199','--call-graph','fp','-o',data,'--',*base,'20000','100'],out/(role+'-profile-probe.json'))
 run(['perf','report','--stdio','--no-children','--no-inline','-i',data],out/(role+'-report.txt'))
 run(['perf','script','--no-inline','-i',data],out/(role+'-perf-script.txt'))
 run(['flamegraph','--perfdata',data,'--no-inline','-o',str(out/(role+'-flamegraph.svg'))],out/(role+'-flamegraph.log'))
 folded=out/'stacks.folded'
 if folded.exists(): folded.rename(out/(role+'-stacks.folded'))
 run(['zstd','-q','-f',data,'-o',str(out/(role+'-perf.data.zst'))])
 run(['zstd','-q','-f','--rm',str(out/(role+'-perf-script.txt'))])
 heap='/tmp/litchi-goal-0416-'+role+'-heaptrack'
 run(['taskset','-c','2','heaptrack','-o',heap,*base,'100','10'],out/(role+'-heap-probe.log'))
 matches=list(pathlib.Path('/tmp').glob(pathlib.Path(heap).name+'*'))
 trace=[p for p in matches if p.suffix in ('.gz','.zst')][0]
 run(['heaptrack_print',str(trace)],out/(role+'-heap-summary.txt'))
 import shutil
 shutil.copy2(trace,out/(role+'-heaptrack'+trace.suffix))
