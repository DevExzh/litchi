#!/usr/bin/env python3
"""Reject concrete corruptions in private copies of the 0412 evidence."""
import argparse,hashlib,json,shutil,subprocess,tempfile
from pathlib import Path

def load(p):return json.loads(p.read_text() if p.exists() else subprocess.check_output(['zstd','-q','-dc',str(p)+'.zst']))
def write(p,v):
 for q in [p,Path(str(p)+'.zst')]:
  if q.exists() or q.is_symlink():q.unlink()
 p.write_text(json.dumps(v,indent=2)+'\n')
def mutate(root,kind):
 p=root/'normal-1.json';v=load(p)
 if kind=='duplicate-order':v['results'][0]['elapsed_ns']['sample_order'][0]=v['results'][0]['elapsed_ns']['sample_order'][1]
 elif kind=='duplicate-metric-index':v['results'][0]['operation_metrics']['sample_indices'][0]=v['results'][0]['operation_metrics']['sample_indices'][1]
 elif kind=='opaque-read':v['results'][1]['source']['xls']['opaque_payload_read_bytes'][0]=1
 elif kind=='wrong-oracle':
  for row in v['results']:row['output_sha256']='0'*64
 elif kind=='missing-owned':
  for q in list(root.glob('owned-[1-4].json*')):
   if '.catalog.' not in q.name:q.unlink()
  return
 elif kind=='bad-binary-identity':
  p=root/'candidate-build-identity.json';v=load(p);v['binaries']['litchi-perf-baseline']['sha256']='0'*64
 write(p,v)
def main():
 p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--verifier',type=Path,required=True);p.add_argument('--resources',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();records=[]
 before={q.name:hashlib.sha256(q.read_bytes()).hexdigest() for q in a.root.iterdir() if q.is_file()}
 for kind,marker,script in [('duplicate-order','sample_order',a.verifier),('duplicate-metric-index','sample_indices',a.verifier),('opaque-read','opaque_payload_read_bytes',a.verifier),('wrong-oracle','output_sha256',a.verifier),('missing-owned','owned',a.verifier),('bad-binary-identity','binary',a.resources)]:
  with tempfile.TemporaryDirectory(prefix='litchi-goal-0412-negative-') as td:
   root=Path(td)
   for q in a.root.iterdir():
    if q.is_file():
     if q.suffix=='.zst':shutil.copy2(q,root/q.name)
     else:(root/q.name).symlink_to(q.resolve())
   mutate(root,kind);argv=['python3',str(script.resolve()),'--root',str(root),'--output',str(root/'probe-result.json')];r=subprocess.run(argv,text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
   if r.returncode==0 or marker.lower() not in r.stdout.lower():raise RuntimeError(f'{kind}: expected rejection containing {marker!r}, got {r.returncode}: {r.stdout}')
   records.append(dict(mutation=kind,argv=argv,exit_code=r.returncode,expected_diagnostic_marker=marker,diagnostic=r.stdout))
 after={q.name:hashlib.sha256(q.read_bytes()).hexdigest() for q in a.root.iterdir() if q.is_file()}
 if before!=after:raise RuntimeError('Original capture changed during negative probes')
 a.output.write_text(json.dumps(dict(status='passed',original_capture_unchanged=True,probes=records),indent=2)+'\n');print(f'Passed {len(records)} expected corruption rejections')
if __name__=='__main__':main()
