#!/usr/bin/env python3
"""Serial captures after an exact bound pilot admission; offline decoding follows."""
import sys,json,subprocess,datetime,hashlib
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0962')
def descriptor(p):
 return {'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def write(p,v):
 with p.open('x') as f:json.dump(v,f,indent=2);f.write('\n')
def main():
 plan_path=Path(sys.argv[1]);plan=json.loads(plan_path.read_text())
 for key in ['capture_plan','pilot_admission','capture_helper','analyzer','wave_helper']:
  p=Path(plan[key]['path']);assert descriptor(p)==plan[key],key
 assert Path(plan['wave_helper']['path']).resolve()==Path(__file__).resolve()
 assert json.loads(Path(plan['pilot_admission']['path']).read_text())['status']=='pass'
 base=json.loads(Path(plan['capture_plan']['path']).read_text())
 root=ROOT/'fp-wave-v1';root.mkdir(exist_ok=False)
 write(root/'start.json',{'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'plan':descriptor(plan_path)})
 rows=[]
 for i,case in enumerate(base['wave_schedule']):
  out=root/f'{i:02d}-{case}'
  cmd=[sys.executable,plan['capture_helper']['path'],'--plan',plan['capture_plan']['path'],'--case',case,'--output',str(out)]
  rc=subprocess.run(cmd).returncode
  row={'index':i,'case':case,'command':cmd,'returncode':rc};rows.append(row)
  write(root/f'child-{i:02d}.json',row)
  if rc:
   write(root/'capture-terminal.json',{'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'status':'failed','children':rows,'planned_not_run':base['wave_schedule'][i+1:]})
   raise RuntimeError('capture failed; wave halted, no replacement')
 write(root/'capture-terminal.json',{'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'status':'capture_validated','children':rows,'attribution_status':'pending offline decode; no wave admission claim'})
if __name__=='__main__':main()
