#!/usr/bin/env python3
"""Run the fixed applicable gate matrix serially under the coordinator lock."""
import subprocess
import sys
from common import ROOT, REPO, ENV

GATES = {
    'libraries-tests-opc-guard': ['cargo', 'test', '--release', '--locked', '-p', 'soapberry-zip', '-p', 'litchi-opc', '-p', 'litchi-pptx'],
    'libraries-clippy-opc-guard': ['cargo', 'clippy', '--release', '--locked', '-p', 'soapberry-zip', '-p', 'litchi-opc', '-p', 'litchi-pptx', '--all-targets', '--no-deps', '--', '-D', 'warnings'],
    'libraries-rustdoc-opc-guard': ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', '--locked', '-p', 'soapberry-zip', '-p', 'litchi-opc', '-p', 'litchi-pptx', '--no-deps'],
    'libraries-format-opc-guard': ['cargo', 'fmt', '-p', 'soapberry-zip', '-p', 'litchi-opc', '-p', 'litchi-pptx', '--', '--check'],
    'shared-streaming-unit': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-docx', '-p', 'litchi-xlsx', '-p', 'litchi-odf-common', '-p', 'litchi-odt', '-p', 'litchi-ods', '-p', 'litchi-odp', '--lib', 'streaming'],
    'docx-streaming': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-docx', '--test', 'streaming'],
    'odf-streaming': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odf-common', '--test', 'streaming_package_writer'],
    'odt-streaming': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odt', '--test', 'streaming_plain_paragraphs'],
    'ods-streaming': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-ods', '--test', 'streaming_creation', '--test', 'streaming_text_spans'],
    'odp-streaming': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odp', '--test', 'streaming_provider'],
    'harness-tests': ['cargo', 'test', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--bin', 'litchi-perf-baseline-alloc', '--bin', 'pptx_metadata_spool'],
    'harness-clippy-opc-guard': ['cargo', 'clippy', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--bin', 'litchi-perf-baseline-alloc', '--bin', 'pptx_metadata_spool', '--no-deps', '--', '-D', 'warnings'],
    'harness-format-opc-guard': ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--', '--check'],
    'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
    'registry-strict': ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict'],
}

if __name__ == '__main__':
    selected = sys.argv[1:] or list(GATES)
    for label in selected:
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *GATES[label]], cwd=REPO, env=ENV, check=True)
