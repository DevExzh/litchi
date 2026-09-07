#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
manifest='/tmp/litchi-goal-0454-opc-fuzz/Cargo.toml'
flags='-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 -C llvm-args=-sanitizer-coverage-inline-8bit-counters -C llvm-args=-sanitizer-coverage-pc-table -C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing'
rows=[('fuzz-lock-r2',['cargo','generate-lockfile','--offline','--manifest-path',manifest]),
      ('fuzz-build-r2',['env','RUSTC_BOOTSTRAP=1','RUSTFLAGS='+flags,'cargo','build','--release','--locked','--manifest-path',manifest,'--target','x86_64-unknown-linux-gnu','--bin','parse_opc']),
      ('fuzz-smoke',['/tmp/litchi-goal-0454-opc-fuzz/target/x86_64-unknown-linux-gnu/release/parse_opc','/tmp/litchi-goal-0454-opc-fuzz/corpus','-runs=1000','-seed=454','-max_len=1048576','-timeout=10','-artifact_prefix=/tmp/litchi-goal-0454-opc-fuzz/']),
      ('final-fuzz-strict',['cargo','clippy','--locked','--release','--manifest-path',manifest,'--','-D','warnings']),
      ('final-fuzz-format',['rustfmt','+1.98.1','--edition','2024','--check','--config','skip_children=true','crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'])]
for tag,argv in rows:
    r=subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*argv])
    if r.returncode:raise SystemExit(r.returncode)
