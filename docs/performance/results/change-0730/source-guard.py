#!/usr/bin/env python3
"""Verify both source states and the live selected disposition against archives."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def read(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
 a=read(P/'baseline-builds.json')['source_sha256'];b=read(P/'candidate-builds.json')['source_sha256'];delta={k for k in set(a)|set(b) if a.get(k)!=b.get(k)}
 record=read(P/'candidate-source.json');assert set(record['files'])==delta
 for rel,r in record['files'].items():
  assert r=={'before_sha256':a.get(rel),'after_sha256':b.get(rel)}
  for variant,expected in [('baseline',a.get(rel)),('candidate',b.get(rel))]:
   if expected is not None:assert sha(P/(variant+'-source')/rel)==expected
  if a.get(rel):assert hashlib.sha256(subprocess.check_output(['git','show',record['base_head']+':'+rel],cwd=ROOT)).hexdigest()==a[rel]
 disposition=read(P/'disposition.json')['production'] if (P/'disposition.json').exists() else 'candidate'
 assert disposition in ['candidate','baseline'];selected=b if disposition=='candidate' else a
 for rel,h in selected.items():assert sha(ROOT/rel)==h,rel
 for rel in (set(a)|set(b))-set(selected):assert not (ROOT/rel).exists(),rel
 for rel,h in read(P/'constraints.json').items():assert sha(ROOT/rel)==h
 assert read(P/'baseline-builds.json')['probe_sha256']==read(P/'candidate-builds.json')['probe_sha256']
 print('PASS baseline/candidate source custody and live '+disposition)
if __name__=='__main__':main()
