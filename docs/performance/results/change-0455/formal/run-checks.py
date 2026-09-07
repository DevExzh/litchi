#!/usr/bin/env python3
"""Serialized release gates for the changed ZIP owner and its consumers."""
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
COMMANDS={
'zip':['cargo','test','--locked','--release','-p','soapberry-zip','--all-features','--','--test-threads=1'],
'opc':['cargo','test','--locked','--release','-p','litchi-opc','--all-features','--','--test-threads=1'],
'pptx':['cargo','test','--locked','--release','-p','litchi-pptx','--all-features','--','--test-threads=1'],
'harness':['cargo','test','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-features','--','--test-threads=1'],
'strict':['cargo','clippy','--locked','--release','-p','soapberry-zip','-p','litchi-opc','-p','litchi-pptx','--all-targets','--all-features','--','-D','warnings'],
'harness-strict':['cargo','clippy','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--all-targets','--all-features','--','-D','warnings'],
'doc':['cargo','doc','--locked','--release','-p','soapberry-zip','-p','litchi-opc','-p','litchi-pptx','--no-deps','--all-features'],
'format':['rustfmt','+1.98.1','--edition','2024','--check','--config','skip_children=true','crates/soapberry-zip/src/preserve.rs','crates/soapberry-zip/tests/preservation_zip64_promotion.rs','tools/perf-baseline/src/pptx_provider_lifecycle.rs'],
'workspace':['cargo','check','--locked','--release','--workspace','--all-targets','--no-default-features','--exclude','litchi-iwa*','--exclude','litchi-keynote','--exclude','litchi-numbers*','--exclude','litchi-pages','--features','litchi/odf'],
'boundaries':['python3','-B','tools/check_crate_boundaries.py'],
}
if __name__=='__main__':
    names=sys.argv[1:] or list(COMMANDS)
    for tag in names:subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*COMMANDS[tag]],check=True)
