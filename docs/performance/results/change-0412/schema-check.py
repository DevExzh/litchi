#!/usr/bin/env python3
"""Replay repository schema and corpus-binding checks for every 0412 capture."""
import argparse,json,subprocess,sys,tempfile
from pathlib import Path

def load(path):
 if path.exists():return json.loads(path.read_text())
 return json.loads(subprocess.check_output(['zstd','-q','-dc',str(path)+'.zst']))

def main():
 p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();sys.path.insert(0,str(a.repo_root.resolve()))
 from tools import perf_compare,validate_perf_corpus_binding
 labels=[f'{family}-{i}' for family,n in [('normal',4),('allocator',4),('owned',4),('owned-allocator',2),('owned-stat',3)] for i in range(1,n+1)]+['owned-profile']
 results=[]
 for label in labels:
  r=load(a.root/(label+'.json'));c=load(a.root/(label+'.catalog.json'))
  with tempfile.TemporaryDirectory(prefix='litchi-goal-0412-schema-') as td:
   rp=Path(td)/'report.json';cp=Path(td)/'catalog.json';rp.write_text(json.dumps(r));cp.write_text(json.dumps(c));validate_perf_corpus_binding.validate_paths(rp,cp)
  perf_compare.validate_parallel_metrics(r)
  count=0
  for v in r['results']:
   elapsed=v['elapsed_ns'];metrics=v['operation_metrics'];perf_compare._validate_operation_metrics(metrics,label,elapsed['samples'],r['schema_version'],elapsed_sample_order=elapsed['sample_order']);count+=len(elapsed['samples'])
  results.append(dict(label=label,rows=len(r['results']),observations=count,status='passed'))
 a.output.write_text(json.dumps(dict(schema_version=1,status='passed',reports=results),indent=2)+'\n');print(f'Passed schema/binding checks for {len(results)} reports')

if __name__=='__main__':main()
