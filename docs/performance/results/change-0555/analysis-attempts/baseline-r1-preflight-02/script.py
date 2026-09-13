import sys,json
from pathlib import Path
b=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(b))
import analyze_profiles as p
import analyze_metrics as m
plan=p.plan_data()
for job in p.profile_jobs(plan):
 if job['repeat']==1:
  row=p.analyze_profile('baseline',job,plan)
  print(job['name'],'validated',len(row['timed_dumps']),'timed dumps',flush=True)
for kind in ('normal','alloc'):
 row=m.binary_metadata('baseline',kind,m.plan_data())
 print('baseline',kind,'binary validated',flush=True)
