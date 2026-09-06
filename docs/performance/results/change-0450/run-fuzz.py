#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
manifest='/tmp/litchi-goal-0450-zip-fuzz/crate/fuzz/Cargo.toml'
flags='-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 -C llvm-args=-sanitizer-coverage-inline-8bit-counters -C llvm-args=-sanitizer-coverage-pc-table -C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing'
rows=[('fuzz-lock',['cargo','generate-lockfile','--offline','--manifest-path',manifest]),
      ('fuzz-build',['env','RUSTC_BOOTSTRAP=1','RUSTFLAGS='+flags,'cargo','build','--release','--locked','--manifest-path',manifest,'--target','x86_64-unknown-linux-gnu','--bin','parse_zip']),
      ('fuzz-smoke',['/tmp/litchi-goal-0450-zip-fuzz/crate/fuzz/target/x86_64-unknown-linux-gnu/release/parse_zip','/tmp/litchi-goal-0450-zip-fuzz/corpus','-runs=1000','-seed=450','-max_len=1048576','-timeout=10','-artifact_prefix=/tmp/litchi-goal-0450-zip-fuzz/']),
      ('final-fuzz-format',['rustfmt','+1.98.1','--edition','2024','--check','--config','skip_children=true','crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs'])]
for tag,argv in rows:
    r=subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*argv])
    if r.returncode:raise SystemExit(r.returncode)
