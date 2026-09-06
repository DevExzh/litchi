#!/usr/bin/env python3
"""Run retained candidate lint and performance build serially after final gates."""
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
p=subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','candidate-harness-strict','--','cargo','clippy','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--all-targets','--no-deps','--','-D','warnings'])
assert p.returncode==1, 'Expected the disclosed inherited harness lint failure; inspect changed outcome'
subprocess.run([sys.executable,'-B',str(ROOT/'lint-comparison.py')],check=True)
subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','after-build','--','env','RUSTFLAGS=-Cforce-frame-pointers=yes','CARGO_PROFILE_RELEASE_DEBUG=1','cargo','build','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc'],check=True)
subprocess.run([sys.executable,'-B',str(ROOT/'save-binaries.py'),'--role','after'],check=True)
subprocess.run([sys.executable,'-B',str(ROOT/'build-descriptor.py'),'--role','after'],check=True)
subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','pilot-after','--',sys.executable,'-B',str(ROOT/'pilot.py'),'--role','after'],check=True)
