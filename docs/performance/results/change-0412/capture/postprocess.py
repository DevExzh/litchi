import datetime,json,os,subprocess,time
from pathlib import Path
out=Path('/tmp/litchi-goal-0412-capture');env=os.environ.copy();env['DEBUGINFOD_URLS']='';records=[]
commands=[('owned-profile-script',['perf','script','--no-inline','-i',str(out/'owned-profile.data')]),('owned-profile-self',['perf','report','--stdio','--no-inline','--no-children','--call-graph=none','--percent-limit=0','-i',str(out/'owned-profile.data')]),('owned-profile-header',['perf','report','--header-only','--stdio','-i',str(out/'owned-profile.data')]),('owned-profile-flamegraph',['flamegraph','--perfdata',str(out/'owned-profile.data'),'--no-inline','--deterministic','--title','Plain OwnedSource XLS open and one cell: whole process','--output',str(out/'owned-profile.svg')])]
for label,argv in commands:
 stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/(label+'.stdout')).open('wb') as stdout,(out/(label+'.stderr')).open('wb') as stderr:r=subprocess.run(argv,env=env,stdout=stdout,stderr=stderr)
 records.append(dict(label=label,argv=argv,started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode));(out/'postprocess-commands.json').write_text(json.dumps(records,indent=2)+'\n');print(label,r.returncode,flush=True);r.check_returncode()
