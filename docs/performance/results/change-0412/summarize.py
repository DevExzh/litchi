#!/usr/bin/env python3
"""Per-repeat uncertainty and bounded review triggers; no production claim."""
import argparse,functools,json,math,subprocess
from pathlib import Path

def load(p):
 return json.loads(p.read_text() if p.exists() else subprocess.check_output(['zstd','-q','-dc',str(p)+'.zst']))
@functools.cache
def median_interval_rank(n):
 cumulative=0;k=0;accepted=0
 for candidate in range(1,n//2+1):
  cumulative+=math.comb(n,candidate-1)
  if cumulative/(2**n)<=.025:k=candidate;accepted=cumulative
  else:break
 return k,1-2*accepted/(2**n)
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();reports={};cis={}
 for family,n in [('normal',4),('owned',4)]:
  for i in range(1,n+1):
   label=f'{family}-{i}';r=load(a.root/(label+'.json'));reports[label]={row['case']:row['elapsed_ns'] for row in r['results']};cis[label]={}
   for row in r['results']:
    values=sorted(row['elapsed_ns']['samples']);k,coverage=median_interval_rank(len(values));cis[label][row['case']]=dict(interval_ns=[values[k-1],values[len(values)-k]],order_statistic=k,iid_binomial_coverage=coverage,samples=len(values))
 guards=[]
 for case in ['xls_semantic_open','xls_eager_open_list_worksheets','xls_eager_open_one_cell']:
  for control,candidate in [('normal-1','normal-2'),('normal-4','normal-3')]:
   for metric in ['p50','p95','p99','mean']:
    delta=(reports[candidate][case][metric]/reports[control][case][metric]-1)*100;guards.append(dict(case=case,control=control,candidate=candidate,metric=metric,candidate_minus_control_percent=delta,exceeds_five_percent_review_trigger=delta>5))
 result=dict(schema_version=1,performance_claim='none',median_intervals=cis,eager_review=guards,limitations=['Exact IID binomial order-statistic intervals describe within-child sample uncertainty only.','They exclude host drift, shared heap/cache dependence and between-process uncertainty; use retained repeat spread too.','Observer versions differ: no production before/after claim.','Review triggers are descriptive and do not prove equivalence.'])
 a.output.write_text(json.dumps(result,indent=2)+'\n');print('Wrote within-child median intervals and eager review triggers')
if __name__=='__main__':main()
