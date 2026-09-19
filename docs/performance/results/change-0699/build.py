#!/usr/bin/env python3
"""Build baseline and rejected codecs only in a disposable sparse worktree."""
import hashlib,json,os,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
WORK=ROOT.parent/'litchi-0699-work'
TARGET=ROOT.parent/'litchi-target-0699'
BIN=ROOT.parent/'litchi-0699-bin'
CODEC='crates/litchi-ooxml-common/src/mce/codec.rs'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def dump(name,data):(P/name).write_text(json.dumps(data,indent=2)+'\n')
phase=sys.argv[1]
assert phase in ['baseline','candidate']
base=json.loads((P/'baseline.json').read_text())
for group in ['constraints_sha256','build_inputs_sha256']:
 for name,digest in base[group].items():assert sha(ROOT/name)==digest,name
assert subprocess.check_output(['git','rev-parse','HEAD'],cwd=WORK,text=True).strip()==base['baseline_head']
assert all(sha(ROOT/n)==h for n,h in base['source_sha256'].items())
if phase=='baseline':
 assert all(sha(WORK/n)==h for n,h in base['source_sha256'].items())
 shutil.copytree(P/'probe',WORK/P.relative_to(ROOT)/'probe',dirs_exist_ok=True)
 shutil.copy2(ROOT/'Cargo.lock',WORK/'Cargo.lock')
else:
 assert all(sha(WORK/n)==h for n,h in base['source_sha256'].items())
 witness=P.parent/'change-0698/candidate-codec.rs.txt'
 assert sha(witness)==base['candidate_codec_sha256']
 (WORK/CODEC).write_bytes(witness.read_bytes())
probe={str(f.relative_to(P)):sha(f) for f in (P/'probe').rglob('*') if f.is_file()}
work_probe=WORK/P.relative_to(ROOT)/'probe'
assert {str(f.relative_to(work_probe)) for f in work_probe.rglob('*') if f.is_file()}=={str(f.relative_to(P/'probe')) for f in (P/'probe').rglob('*') if f.is_file()}
assert all(sha(WORK/P.relative_to(ROOT)/n)==h for n,h in probe.items())
assert all(sha(WORK/n)==h for n,h in base['build_inputs_sha256'].items())
if phase=='candidate':assert probe==json.loads((P/'build-baseline.json').read_text())['probe_sha256']
source={n:sha(WORK/n) for n in base['source_sha256']}
if phase=='candidate':assert {n for n in source if source[n]!=base['source_sha256'][n]}=={CODEC}
manifest=WORK/P.relative_to(ROOT)/'probe/Cargo.toml'
command=['cargo','build','--release','--locked','--manifest-path',str(manifest),'--target-dir',str(TARGET),'-j','2']
start=time.monotonic()
with (P/f'build-{phase}.log').open('w') as out:
 result=subprocess.run(command,cwd=WORK,env={**os.environ,'RUSTFLAGS':'-D warnings'},stdout=out,stderr=subprocess.STDOUT)
assert result.returncode==0
assert all(sha(WORK/n)==h for n,h in source.items())
assert all(sha(P/n)==h for n,h in probe.items())
BIN.mkdir(exist_ok=True)
binary=BIN/phase
shutil.copy2(TARGET/'release/probe0699-refusal',binary)
dump(f'build-{phase}.json',dict(phase=phase,command=command,exit_code=0,seconds=time.monotonic()-start,source_sha256=source,probe_sha256=probe,binary=str(binary),binary_sha256=sha(binary),log_sha256=sha(P/f'build-{phase}.log')))
if phase=='candidate':
 (WORK/CODEC).write_bytes((ROOT/CODEC).read_bytes())
 assert all(sha(WORK/n)==h for n,h in base['source_sha256'].items())
 dump('build-state-verification.json',dict(worktree_build_inputs_sha256={n:sha(WORK/n) for n in base['build_inputs_sha256']},worktree_restored_source_sha256={n:sha(WORK/n) for n in base['source_sha256']},probe_sha256=probe,main_source_unchanged=all(sha(ROOT/n)==h for n,h in base['source_sha256'].items())))
print(phase,'built; main production unchanged',flush=True)
