#!/usr/bin/env python3
"""Run the unchanged OPC publication fuzz target against the candidate ZIP owner."""
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
TASK='/tmp/litchi-goal-0455/fuzz'
FLAGS='-C passes=sancov-module -C llvm-args=-sanitizer-coverage-level=4 -C llvm-args=-sanitizer-coverage-inline-8bit-counters -C llvm-args=-sanitizer-coverage-pc-table -C llvm-args=-sanitizer-coverage-trace-compares -Z sanitizer=address --cfg fuzzing'
COMMANDS={
'fuzz-lock':['cargo','generate-lockfile','--offline','--manifest-path',TASK+'/Cargo.toml'],
'fuzz-build':['env','RUSTC_BOOTSTRAP=1','RUSTFLAGS='+FLAGS,'cargo','build','--release','--locked','--manifest-path',TASK+'/Cargo.toml','--target','x86_64-unknown-linux-gnu','--bin','parse_opc'],
'fuzz-smoke':[TASK+'/target/x86_64-unknown-linux-gnu/release/parse_opc',TASK+'/corpus','-runs=1000','-seed=455','-max_len=1048576','-timeout=10','-artifact_prefix='+TASK+'/'],
}
if __name__=='__main__':
    for tag,argv in COMMANDS.items():subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*argv],check=True)
