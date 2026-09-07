#!/usr/bin/env python3
import json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
commands=[
 ('final-format-r2',['rustfmt','+1.98.1','--edition','2024','--check','--config','skip_children=true',*json.loads((ROOT/'source-files.json').read_text())]),
 ('final-strict-r7',['cargo','clippy','--locked','--release','-p','litchi-pptx','-p','litchi-opc','--all-targets','--all-features','--','-D','warnings']),
 ('final-harness-strict-r4',['cargo','clippy','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-targets','--all-features','--','-D','warnings']),
 ('final-opc-r4',['cargo','test','--locked','--release','-p','litchi-opc','--all-features','--','--test-threads=1']),
 ('final-pptx-r4',['cargo','test','--locked','--release','-p','litchi-pptx','--all-features','--','--test-threads=1']),
 ('final-harness-r3',['cargo','test','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--','--test-threads=1']),
 ('final-doc-r2',['cargo','doc','--locked','--release','-p','litchi-pptx','-p','litchi-opc','--no-deps','--all-features']),
 ('final-workspace-r2',['cargo','check','--locked','--release','--workspace','--all-targets','--no-default-features','--exclude','litchi-iwa*','--exclude','litchi-keynote','--exclude','litchi-numbers*','--exclude','litchi-pages','--features','litchi/odf']),
 ('final-boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
for tag,args in commands:subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*args],check=True)
