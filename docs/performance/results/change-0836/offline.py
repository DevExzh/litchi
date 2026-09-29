"""Retain root-run offline validation outcomes without executing workloads."""
import subprocess,sys,time
import driver as d

def run(label,argv):
 out=d.P/'validation'/label;out.mkdir(parents=True,exist_ok=False)
 start=time.time();d.write(out/'started.json',dict(argv=argv,cwd=str(d.ROOT),started_unix=start))
 with (out/'output.log').open('xb') as f:
  code=subprocess.run(argv,cwd=d.ROOT,env=d.env(),stdout=f,stderr=subprocess.STDOUT).returncode
 d.write(out/'receipt.json',dict(argv=argv,cwd=str(d.ROOT),exit_code=code,started_unix=start,finished_unix=time.time(),log=d.desc(out/'output.log')))
 print(label,code,flush=True)
 if code:print((out/'output.log').read_text()[-5000:],flush=True)
 return code
if __name__=='__main__':sys.exit(run(sys.argv[1],sys.argv[2:]))
