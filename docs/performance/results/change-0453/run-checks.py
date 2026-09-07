#!/usr/bin/env python3
import json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
commands=[
 ('final-strict',['cargo','clippy','--locked','--release','-p','litchi-pptx','--all-targets','--all-features','--','-D','warnings']),
 ('final-harness-strict',['cargo','clippy','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-targets','--all-features','--','-D','warnings']),
 ('final-pptx',['cargo','test','--locked','--release','-p','litchi-pptx','--all-features','--','--test-threads=1']),
 ('final-opc',['cargo','test','--locked','--release','-p','litchi-opc','--all-features','--','--test-threads=1']),
 ('final-harness',['cargo','test','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--','--test-threads=1']),
 ('final-doc',['cargo','doc','--locked','--release','-p','litchi-pptx','--no-deps','--all-features']),
 ('final-workspace',['cargo','check','--locked','--release','--workspace','--all-targets','--no-default-features','--exclude','litchi-iwa*','--exclude','litchi-keynote','--exclude','litchi-numbers*','--exclude','litchi-pages','--features','litchi/odf']),
 ('final-format',['rustfmt','+1.98.1','--edition','2024','--check','--config','skip_children=true',*json.loads((ROOT/'source-files.json').read_text())]),
 ('final-boundaries',['python3','-B','tools/check_crate_boundaries.py'])]
for tag,args in commands:subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag+'-r2','--',*args],check=True)
subprocess.run([sys.executable,'-B',str(ROOT/'build.py'),'candidate'],check=True)
