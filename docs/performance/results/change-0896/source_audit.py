#!/usr/bin/env python3
"""Audit the complete isolated source census before/after native builds."""
from pathlib import Path
import argparse, hashlib, json
P=Path(__file__).resolve().parent
R=P.parents[3]
S=R/'target/litchi-0896-source'
def sha(f):
 h=hashlib.sha256()
 with f.open('rb') as s:
  for b in iter(lambda:s.read(1024*1024),b''):h.update(b)
 return h.hexdigest()
def d(f):return {'path':str(f.resolve()),'bytes':f.stat().st_size,'sha256':sha(f)}
def main():
 a=argparse.ArgumentParser();a.add_argument('--output',type=Path,required=True);a.add_argument('--before',type=Path);args=a.parse_args()
 base=json.loads((P.parent/'change-0894/source.json').read_text())['files']
 candidate=json.loads((P/'candidate-source.json').read_text())
 support=json.loads((P/'source.json').read_text())['support']
 files={}
 for f in S.rglob('*'):
  if f.is_dir():continue
  assert f.is_file() and not f.is_symlink(),f
  files[f.relative_to(S).as_posix()]={'bytes':f.stat().st_size,'sha256':sha(f)}
 extra={'Cargo.lock','.owner-litchi-0896'}|set(support)
 assert set(files)==set(base)|extra,sorted(set(files)^(set(base)|extra))
 for n,v in base.items():assert files[n]['sha256']==candidate['overlay'].get(n,v)['sha256'],n
 for n,v in support.items():assert files[n]['sha256']==v['sha256'],n
 assert files['Cargo.lock']['sha256']==sha(P/'quality-Cargo.lock')
 result={'status':'pass','baseline_commit':'506aa2edae6666d9c76ae6827c2f5941a8e47896','evaluation_head':'3ff493b638b7aa5860d97858dae8bd82bac0f2d5','schema':'litchi.qname-source-census-0896.v1','helper':d(Path(__file__)),'source':d(P/'source.json'),'candidate_source':d(P/'candidate-source.json'),'selected_file_count':len(base),'overlay_replacements':len(candidate['overlay']),'file_count':len(files),'files':dict(sorted(files.items()))}
 if args.before:
  before=json.loads(args.before.read_text());assert before['files']==result['files'],'build changed frozen source'
  result['before']=d(args.before)
 with args.output.open('x') as f:json.dump(result,f,indent=2,sort_keys=True);f.write('\n')
 print('Source census PASS',len(files))
if __name__=='__main__':main()
