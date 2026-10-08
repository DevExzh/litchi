import sys, json, pathlib, hashlib, os, shutil, subprocess, time, datetime
R=pathlib.Path('/home/zhuhe/litchi-performance-0977'); repo=pathlib.Path('/home/zhuhe/code/litchi')
sys.path.insert(0,str(R/'tools'))
import plan_binding_0977_v2 as c

def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
def put(p,v):
    with p.open('x') as f: json.dump(v,f,indent=2,sort_keys=True); f.write('\n')
def quiet():
    active=[]
    for d in pathlib.Path('/proc').iterdir():
        if not d.name.isdigit() or int(d.name)==os.getpid(): continue
        try: a=[x.decode(errors='replace') for x in (d/'cmdline').read_bytes().split(b'\0') if x]
        except (OSError,ProcessLookupError): continue
        names=[pathlib.Path(x).name for x in a]
        if (names and names[0]=='perf' and 'record' in a) or (names and names[0]=='doc_retention_probe_0942_cfb_profile') or ('run_0977_v2.py' in names and '--run' in a): active.append({'pid':int(d.name),'argv':a})
    assert not active,active
    return active
planpath=R/'plans/plan-0977-v2.json'; plan=json.loads(planpath.read_text())
audit=pathlib.Path(sys.argv[1]); reviewpath=pathlib.Path(sys.argv[2]); review=json.loads(reviewpath.read_text())
assert review['status']=='pass' and review['plan']==c.binding(planpath)
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip()=='ba80617f7fe0025b8a850de459bc303efa4d6290'
c.validate_plan(planpath)
assert not pathlib.Path(plan['outputs']['capture_root']).exists()
active=quiet(); assert {30,31} <= os.sched_getaffinity(0)
free=shutil.disk_usage(R).free; assert free>20*1024**3
manifest=R/'capture/affinity-control-v1/capture-manifest.json'
go={'schema':c.GO_SCHEMA,'authorized':True,'status':'authorized','mode':'office','created_utc':now(),'plan':c.binding(planpath),'helper_seal':plan['helper_seal'],'independent_review':c.binding(reviewpath),'control_audit':c.binding(audit),'control_manifest':c.binding(manifest),'head':'ba80617f7fe0025b8a850de459bc303efa4d6290','quiescence':{'status':'pass','active_capture_or_target_processes':active,'allowed_cpus':sorted(os.sched_getaffinity(0)),'free_bytes':free,'cooldown_seconds':10},'synthetic_boundary':{'enabled':False},'scope':'Exactly12serialOfficechildren;stopfirstfailure;no retry or selective replacement.'}
gopath=R/'receipts/root-go-0977-office-v1.json'; put(gopath,go)
cmd=['/usr/bin/python3.14','-B',str(R/'tools/run_0977_v2.py'),'--run','--plan',str(planpath),'--go',str(gopath)]
env={'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','TZ':'UTC','PYTHONDONTWRITEBYTECODE':'1'}
put(R/'receipts/root-office-command-0977-v1.json',{'argv':cmd,'created_utc':now(),'cwd':str(repo),'environment':env,'go':c.binding(gopath)})
time.sleep(10);quiet();c.validate_plan(planpath)
started=now(); out=R/'receipts/root-office-0977-v1.stdout'; err=R/'receipts/root-office-0977-v1.stderr'
with out.open('x') as o,err.open('x') as e:
    p=subprocess.Popen(cmd,cwd=repo,env=env,stdout=o,stderr=e)
    print(json.dumps({'started':started,'runner_pid':p.pid,'go':c.binding(gopath)}),flush=True)
    rc=p.wait()
terminal={'schema':'litchi.performance.0977.root_terminal.v1','mode':'office','started_utc':started,'finished_utc':now(),'returncode':rc,'runner_pid':p.pid,'command':cmd,'stdout':c.binding(out),'stderr':c.binding(err),'retry':False}
put(R/'receipts/root-office-terminal-0977-v1.json',terminal)
print(json.dumps(terminal),flush=True)
