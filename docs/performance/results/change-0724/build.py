#!/usr/bin/env python3
"""Build four source-bound diagnostic variants serially, then restore baseline."""
import hashlib,json,os,shutil,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];OLD=P.parent/'change-0723'
TARGET=ROOT.parent/'litchi-target-0724';BIN=ROOT.parent/'litchi-0724-bin'
PATHS=['crates/litchi-xls/src/workbook/query_cache.rs','crates/litchi-xls/src/workbook/source.rs']
def sha(b):return hashlib.sha256(b).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def census():return {str(p.relative_to(ROOT)):sha(p.read_bytes()) for c in ('litchi-cfb','litchi-xls') for p in sorted((ROOT/'crates'/c).rglob('*')) if p.is_file() and p.suffix in ('.rs','.toml')}
original={r:(ROOT/r).read_bytes() for r in PATHS}
for r,b in original.items():assert b==subprocess.check_output(['git','show','0d943df447:'+r],cwd=ROOT)
full={r:(OLD/'candidate-source'/r).read_bytes() for r in PATHS}
base_source=original[PATHS[1]].decode();full_source=full[PATHS[1]].decode()
start='    let mut chain = index.chain_checkpoint.as_ref().map_or_else('
end='        if let Some(context) = execution {'
a=base_source.index(start,base_source.index('fn replay_indexed_cell('));b=base_source.index(end,a)
c=full_source.index('    let slots = index.slots_for(row, column);',full_source.index('fn replay_indexed_cell('));d=full_source.index(end,c)
selection_source=(base_source[:a]+full_source[c:d]+base_source[b:]).encode()
variants={'baseline':original,'layout':{PATHS[0]:full[PATHS[0]],PATHS[1]:original[PATHS[1]]},'selection':{PATHS[0]:full[PATHS[0]],PATHS[1]:selection_source},'full':full}
# These attribution variants deliberately leave fields/helpers unused. Permit
# only those items in diagnostic source copies; production and full stay exact.
for variant in ('layout','selection'):
 text=variants[variant][PATHS[0]].decode()
 for method in ('target_slot_and_chain_checkpoint','set_target_chain_checkpoint'):
  text=text.replace('    pub(crate) fn '+method+'(', '    #[allow(dead_code)] // Diagnostic ablation omits checkpoint construction.\n    pub(crate) fn '+method+'(')
 if variant=='layout':
  text=text.replace('    pub(crate) target_chain_checkpoint:', '    #[allow(dead_code)] // Diagnostic layout-only field.\n    pub(crate) target_chain_checkpoint:')
 variants[variant][PATHS[0]]=text.encode()
assert not (P/'builds.json').exists()
rows=[];baseline=census();write(P/'source-baseline.json',baseline)
try:
 for variant,files in variants.items():
  for rel,data in files.items():
   (ROOT/rel).write_bytes(data);archive=P/'sources'/variant/rel;archive.parent.mkdir(parents=True,exist_ok=True);archive.write_bytes(data)
  sources=census();write(P/('source-'+variant+'.json'),sources)
  for folder,binary in [('change-0684/repeat-probe','xls0684-repeat'),('change-0686/probe','xls-index-retry-probe-0686')]:
   dest=BIN/variant/binary;dest.parent.mkdir(parents=True,exist_ok=True)
   cmd=['cargo','build','--manifest-path','docs/performance/results/'+folder+'/Cargo.toml','--release','--locked','--offline'];t=time.monotonic()
   log=P/(variant+'-'+binary+'.build.log')
   with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_TARGET_DIR=str(TARGET),CARGO_BUILD_JOBS='2'),stdout=f,stderr=subprocess.STDOUT)
   assert sources==census(),'source mutated during build'
   row=dict(variant=variant,command=cmd,exit_code=r.returncode,seconds=time.monotonic()-t,source_manifest='source-'+variant+'.json',log=log.name)
   if r.returncode==0:
    shutil.copy2(TARGET/'release'/binary,dest);row.update(binary=str(dest),binary_sha256=sha(dest.read_bytes()),bytes=dest.stat().st_size)
   rows.append(row);write(P/'builds.json',rows);assert r.returncode==0,row
   print(variant,binary,'complete',flush=True)
finally:
 for rel,data in original.items():(ROOT/rel).write_bytes(data)
 assert baseline==census();write(P/'build-restoration.json',dict(exact=True,source_sha256=baseline))
