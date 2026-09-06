#!/usr/bin/env python3
"""Run applicable final gates serially against one unchanged production candidate."""
import json, subprocess, sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
commands=[
 ('candidate-opc-tests',['cargo','test','--locked','--release','-p','litchi-opc','--all-features','--','--test-threads=1']),
 ('candidate-opc-strict',['cargo','clippy','--locked','--release','-p','litchi-opc','--all-features','--all-targets','--no-deps','--','-D','warnings']),
 ('candidate-opc-doc',['cargo','doc','--locked','--release','-p','litchi-opc','--all-features','--no-deps']),
 ('candidate-harness-tests',['cargo','test','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--','--test-threads=1']),
 ('candidate-workspace-check',['cargo','check','--locked','--release','--workspace','--all-targets','--no-default-features','--exclude','litchi-iwa*','--exclude','litchi-keynote','--exclude','litchi-numbers*','--exclude','litchi-pages','--features','litchi/odf']),
 ('candidate-harness-check',['cargo','check','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--all-targets']),
 ('candidate-harness-doc',['cargo','doc','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--no-deps']),
 ('final-format',['rustfmt','--check','--edition','2024','--config','skip_children=true']+json.loads((ROOT/'source-files.json').read_text())),
 ('final-boundaries',['python3','-B','tools/check_crate_boundaries.py']),
]
for tag,command in commands:
    print(tag,flush=True)
    subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--']+command,check=True)
