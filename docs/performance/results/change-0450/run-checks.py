#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
commands=[
 ('final-zip-tests',['cargo','test','--locked','--release','-p','soapberry-zip','--','--test-threads=1','--nocapture']),
 ('final-opc-tests',['cargo','test','--locked','--release','-p','litchi-opc','--','--test-threads=1']),
 ('final-pptx-cross-copy-tests',['cargo','test','--locked','--release','-p','litchi-pptx','--test','source_backed_cross_copy','--test','source_backed_cross_copy_adversarial','--','--test-threads=1']),
 ('final-zip-check',['cargo','check','--locked','--release','-p','soapberry-zip','--all-targets','--all-features']),
 ('final-zip-doc',['cargo','doc','--locked','--release','-p','soapberry-zip','--no-deps','--all-features']),
 ('final-workspace-check',['cargo','check','--locked','--release','--workspace','--all-targets','--no-default-features','--exclude','litchi-iwa*','--exclude','litchi-keynote','--exclude','litchi-numbers*','--exclude','litchi-pages','--features','litchi/odf']),
 ('final-format',['rustfmt','+1.98.1','--edition','2021','--check','--config','skip_children=true','crates/soapberry-zip/src/office.rs']),
 ('final-boundaries',['python3','-B','tools/check_crate_boundaries.py']),
]
for tag,argv in commands:
    result=subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*argv])
    if result.returncode:raise SystemExit(result.returncode)
