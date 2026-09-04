#!/usr/bin/env python3
"""Publish captured evidence with a self-contained hash verifier."""
import hashlib,json,shutil,subprocess
from pathlib import Path
repo=Path.cwd();dest=repo/'docs/performance/results/change-0409';dest.mkdir(parents=True)
originals=[]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for source,subdir in [('capture','initial'),('corrected','corrected'),('attribution','attribution'),('pmu','pmu')]:
 src=Path('/tmp/litchi-goal-0409-'+source);out=dest/subdir;out.mkdir()
 for p in sorted(src.iterdir()):
  if not p.is_file() or p.name in ['cache_touch','review-eba1f-diff.txt','review-eba1f-stat.txt']:continue
  if p.name in ['perf-script.stdout','perf-script-no-inline.stdout']:
   compressed=Path(str(p)+'.zst')
   if not compressed.exists():subprocess.run(['zstd','-q','-f',str(p),'-o',str(compressed)],check=True)
   shutil.copy2(compressed,out/compressed.name)
   originals.append({'path':f'{subdir}/{p.name}','compressed_path':f'{subdir}/{compressed.name}','bytes':p.stat().st_size,'sha256':sha(p)})
  else:shutil.copy2(p,out/p.name)
checks=dest/'checks';checks.mkdir()
for p in sorted(Path('/tmp').glob('litchi-goal-0409-*')):
 if p.is_file() and (p.suffix=='.log' or p.name.endswith('build-commands.txt')):shutil.copy2(p,checks/p.name.replace('litchi-goal-0409-',''))
shutil.copy2(__file__,dest/'publish.py')
verifier=(repo/'docs/performance/results/change-0408/verify-artifacts.py').read_text().replace('0408','0409')
(dest/'verify-artifacts.py').write_text(verifier)
(dest/'README.md').write_text('''# Change 0409 evidence

This is descriptive XLSX evidence, not an accepted before/after claim.

- `initial/`: source abe38a9570129c6646bb1b1d7207c407fc86c3d6; seven reports and their exact command journal, CPU profile, flamegraph, catalogs and hashes. The source-edit workbook/worksheet counters are unconfigured false zeroes; total source, timing, sink and preservation evidence remain valid.
- `corrected/`: committed harness range correction; source-edit 500 samples plus two-sheet, all-sheet batch, and managed smoke checks. `compressed-member-intersections-v1` identifies cumulative compressed source ranges across open/planning/commit/publication. Positive unselected reads include untouched raw publication, not semantic decoding.
- `attribution/`: reproducible period-weighted stack parser and the all-symbol/no-inline profile; `SelectedWorksheet::cell` identifies a subset of the timed query. Inclusive rows overlap. Sample periods do not measure wall-clock phase duration.
- `pmu/`: local sysfs/perf inventory and controlled native-L2 event probes. Exact LLC events are unavailable and all-zero generic L1 aliases are unusable on this guest. Each probe event was measured in a separate invocation.
- `checks/`: build/test/validation logs, including initial formatting and pinned-toolchain failures with their corrected checks.

The command journals retain original temporary paths. Reproduction requires rebuilding the recorded source with Rust 1.98.1 and the exact recorded environment, then adjusting the task-specific output paths. The capture scripts assert the original source/worktree and deterministic corpus identities. The user's untracked docs/GOAL.md is recorded as the sole dirty entry. No clean ABBA, cold-cache, scaling, full CRUD coverage or production speedup claim is authorized.

Large symbolized perf scripts are retained losslessly as Zstandard. Raw perf.data, symbolized reports, the flamegraph, and binaries' hashes/build IDs remain; the large build executables and PMU probe executable are regenerated from the retained sources/commands. Original capture artifacts.json describes the original temporary capture, including the now-compressed script. The top-level artifact-manifest.json is authoritative for the published layout.

Run `python3 verify-artifacts.py` from any directory to verify every published file and both decompressed scripts. Zstandard is required. The parser in attribution/ can consume a decompressed script to reproduce its JSON summary.
''')
files=[{'path':p.relative_to(dest).as_posix(),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(dest.rglob('*')) if p.is_file() and p.name!='artifact-manifest.json']
(dest/'artifact-manifest.json').write_text(json.dumps({'schema_version':1,'performance_claim':'none','files':files,'compressed_originals':originals},indent=2)+'\n')
subprocess.run(['python3',str(dest/'verify-artifacts.py')],check=True)
print('Published',len(files),'files')
