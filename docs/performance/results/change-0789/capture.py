"""Root-only, serial known-page calibration; preserve raw protocol and accounting."""
import gzip,json,os,sys,time
from pathlib import Path
import common as c


def identity(pid):
 p=Path('/proc')/str(pid);raw=(p/'stat').read_text();tail=raw[raw.rfind(')')+2:].split()
 return {'pid':pid,'ppid':int(tail[1]),'starttime':int(tail[19]),'exe':os.readlink(p/'exe'),'stat':raw,'task_ids':sorted(int(x.name) for x in (p/'task').iterdir())}


def usage(r):return {'maxrss_kib':r.ru_maxrss,'minor_faults':r.ru_minflt,'major_faults':r.ru_majflt,'user_seconds':r.ru_utime,'system_seconds':r.ru_stime}


def run_child(out,index,case,affinity,launcher,observer,repeat,trace=False):
 stem=f'{index:03}';report=out/f'{stem}.json';err=out/f'{stem}.stderr.log';transcript=out/f'{stem}.transcript.txt';points_file=out/f'{stem}.points.json.gz';time_file=out/f'{stem}.time.txt';trace_file=out/f'{stem}.strace.log'
 assert not any(p.exists() for p in [report,err,transcript,points_file,time_file,trace_file])
 binary=c.read(c.P/'build.json')['binary'];c.verify(binary)
 args=[binary['path'],'--mib',str(case['mib']),'--workers',str(case['workers']),'--touch',case['touch']]
 if launcher=='time':args=['/usr/bin/time','-f','%M %R %F','-o',str(time_file),*args]
 if trace:args=['/usr/bin/strace','-f','-qq','-e','trace=execve,getrusage,wait4','-o',str(trace_file),*args]
 cpus=','.join(map(str,c.read(c.P/'plan.json')['affinities'][affinity]));command=['/usr/bin/taskset','-c',cpus,*args]
 in_r,in_w=os.pipe();out_r,out_w=os.pipe();errfd=os.open(err,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
 actions=[(os.POSIX_SPAWN_DUP2,in_r,0),(os.POSIX_SPAWN_DUP2,out_w,1),(os.POSIX_SPAWN_DUP2,errfd,2)]+[(os.POSIX_SPAWN_CLOSE,fd) for fd in [in_r,in_w,out_r,out_w,errfd]]
 env=os.environ.copy();started=time.time();pid=os.posix_spawn(command[0],command,env,file_actions=actions)
 for fd in [in_r,out_w,errfd]:os.close(fd)
 points=[];lines=[];error=None;first=None
 reader=os.fdopen(out_r,'rb');writer=os.fdopen(in_w,'wb',buffering=0)
 try:
  for phase in c.read(c.P/'plan.json')['phases']:
   line=reader.readline();lines.append(line);fields=line.decode('ascii').strip().split('\t');assert len(fields)==8 and fields[:2]==['RSS0789',phase],repr(line)
   child=int(fields[2]);ident=identity(child);assert ident['exe']==binary['path']
   stamp=(child,ident['ppid'],ident['starttime'])
   if first is None:first=stamp
   else:assert stamp==first
   if launcher=='direct':assert child==pid and ident['ppid']==os.getpid()
   else:
    parent=identity(ident['ppid']);assert parent['exe']=='/usr/bin/time'
    if trace:assert parent['ppid']==pid and identity(pid)['exe']=='/usr/bin/strace'
    else:assert ident['ppid']==pid
   assert ident['task_ids']==[child]
   snap={'stat':ident.pop('stat')}
   if observer=='full':
    for name in ['smaps','smaps_rollup','maps','status']:snap[name]=(Path('/proc')/str(child)/name).read_text()
   pre=dict(zip(['maxrss_kib','minor_faults','major_faults','mapped_bytes','checksum'],map(int,fields[3:])))
   writer.write(b'+\n')
   line=reader.readline();lines.append(line);fields=line.decode('ascii').strip().split('\t');assert len(fields)==6 and fields[:3]==['ACK0789',phase,str(child)],repr(line)
   after=dict(zip(['maxrss_kib','minor_faults','major_faults'],map(int,fields[3:])))
   points.append({'phase':phase,'identity':ident,'before':pre,'after':after,'raw':snap})
 except Exception as exc:error=repr(exc)
 finally:writer.close()
 extra=reader.read();reader.close()
 if extra:lines.append(extra);error=error or 'unexpected trailing stdout'
 waited,status,r=os.wait4(pid,0);assert waited==pid
 transcript.write_bytes(b''.join(lines))
 points_file.write_bytes(gzip.compress((json.dumps(points,indent=2,sort_keys=True)+'\n').encode(),mtime=0))
 result={'schema':'litchi.rss-accounting-child.0789.v1','index':index,'case':case,'affinity':affinity,'launcher':launcher,'observer':observer,'repeat':repeat,'trace':trace,'command':command,'binary':binary,'started':started,'ended':time.time(),'spawn_method':'os.posix_spawn taskset exec','wrapper_pid':pid,'wait4':{'pid':waited,'status':status,'exit_code':os.waitstatus_to_exitcode(status),'rusage':usage(r)},'points':c.artifact(points_file),'transcript':c.artifact(transcript),'stderr':c.artifact(err),'error':error}
 if time_file.exists():result['gnu_time']=c.artifact(time_file)
 if trace_file.exists():result['strace']=c.artifact(trace_file)
 c.write(report,result)
 assert error is None and result['wait4']['exit_code']==0,result
 print(out.name,index,case,affinity,launcher,observer,'PASS',flush=True)
 return c.artifact(report)


def main():
 lane=sys.argv[1];assert lane in ['matrix','traces'];out=c.P/lane;assert not out.exists();out.mkdir()
 plan=c.read(c.P/'plan.json');rows=[]
 if lane=='matrix':
  for repeat,order in enumerate(plan['orders']):
   for case in plan['cases']:
    for affinity in plan['affinities']:
     for label in order:
      launcher,observer=label.split('-');rows.append(run_child(out,len(rows),case,affinity,launcher,observer,repeat))
      c.write(out/'receipts.json',rows)
 else:
  for item in plan['traces']:
   case={k:item[k] for k in ['mib','workers','touch']}
   rows.append(run_child(out,len(rows),case,item['affinity'],'time','full',0,trace=True));c.write(out/'receipts.json',rows)
 assert len(rows)==plan['expected'][lane+'_children' if lane=='matrix' else 'trace_children']
 c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json')})

if __name__=='__main__':main()
