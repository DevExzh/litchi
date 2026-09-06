#!/usr/bin/env python3
import subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
commands=[('final-opc-tests', ['cargo', 'test', '--locked', '--release', '-p', 'litchi-opc', '--all-features', '--', '--test-threads=1']), ('final-pptx-tests', ['cargo', 'test', '--locked', '--release', '-p', 'litchi-pptx', '--all-features', '--', '--test-threads=1']), ('final-harness-tests', ['cargo', 'test', '--locked', '--release', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--all-features', '--', '--test-threads=1']), ('candidate-harness-build', ['cargo', 'build', '--locked', '--release', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--bin', 'litchi-perf-baseline']), ('final-doc', ['cargo', 'doc', '--locked', '--release', '-p', 'litchi-opc', '-p', 'litchi-pptx', '--no-deps', '--all-features']), ('final-workspace-check', ['cargo', 'check', '--locked', '--release', '--workspace', '--all-targets', '--no-default-features', '--exclude', 'litchi-iwa*', '--exclude', 'litchi-keynote', '--exclude', 'litchi-numbers*', '--exclude', 'litchi-pages', '--features', 'litchi/odf']), ('final-format', ['rustfmt', '+1.98.1', '--edition', '2024', '--check', '--config', 'skip_children=true', 'crates/litchi-opc/src/source_backed.rs', 'crates/litchi-opc/src/lib.rs', 'crates/litchi-opc/tests/source_part_transfer.rs', 'crates/litchi-pptx/src/presentation/source_cross_copy.rs', 'crates/litchi-pptx/tests/source_backed_cross_copy.rs', 'crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs']), ('final-boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py'])]
commands=[('final-strict', ['cargo','clippy','--locked','--release','-p','litchi-opc','-p','litchi-pptx','--all-targets','--all-features','--','-D','warnings'])]+commands
for tag,argv in commands:
    tag += '-r3'
    r=subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',tag,'--',*argv])
    if r.returncode:raise SystemExit(r.returncode)
