"""Root command custody with separate stdout for large profiler artifacts."""
import subprocess,time
import driver as d

def run(stage,label,argv,stdout=None):
 d.check(stage)
 out=d.P/'commands'/label;out.mkdir(parents=True,exist_ok=False)
 start=time.time();d.write(out/'started.json',dict(argv=argv,cwd=str(d.ROOT),started_unix=start,freeze=d.desc(d.P/f'freeze-{stage}.json'),runner=d.desc(d.P/'runner.py')))
 code=None;error=None
 with (out/'output.log').open('xb') as log:
  stream=log if stdout is None else stdout.open('xb')
  try:
   try:code=subprocess.run(argv,cwd=d.ROOT,env=d.env(stage),stdout=stream,stderr=log).returncode
   except Exception as e:error=repr(e)
  finally:
   if stdout is not None:stream.close()
 row=dict(argv=argv,exit_code=code,error=error,started_unix=start,finished_unix=time.time(),log=d.desc(out/'output.log'),freeze=d.desc(d.P/f'freeze-{stage}.json'),runner=d.desc(d.P/'runner.py'))
 if stdout is not None and stdout.exists():row['stdout']=d.desc(stdout)
 d.write(out/'receipt.json',row);print(label,code,error,flush=True)
 return row
