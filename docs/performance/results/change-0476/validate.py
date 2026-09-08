#!/usr/bin/env python3
"""Run the selected Rust and repository checks serially under the caller's lock."""
import subprocess
import sys
from common import ROOT, REPO, ENV


COMMANDS = [
    ('zip-tests', ['cargo', 'test', '--release', '--locked', '-p', 'soapberry-zip']),
    ('opc-tests', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-opc']),
    ('shared-streaming-tests', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-docx', '-p', 'litchi-pptx', '-p', 'litchi-xlsx', '-p', 'litchi-odf-common', '-p', 'litchi-odt', '-p', 'litchi-ods', '-p', 'litchi-odp', 'streaming']),
    ('shared-streaming-unit', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-docx', '-p', 'litchi-pptx', '-p', 'litchi-xlsx', '-p', 'litchi-odf-common', '-p', 'litchi-odt', '-p', 'litchi-ods', '-p', 'litchi-odp', '--lib', 'streaming']),
    ('docx-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-docx', '--test', 'streaming']),
    ('pptx-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-pptx', '--test', 'streaming_authoring']),
    ('odf-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odf-common', '--test', 'streaming_package_writer']),
    ('odt-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odt', '--test', 'streaming_plain_paragraphs']),
    ('ods-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-ods', '--test', 'streaming_creation', '--test', 'streaming_text_spans']),
    ('odp-streaming', ['cargo', 'test', '--release', '--locked', '-p', 'litchi-odp', '--test', 'streaming_provider']),
    ('zip-focused-final', ['cargo', 'test', '--release', '--locked', '-p', 'soapberry-zip', '--test', 'deflate_reuse']),
    ('harness-tests', ['cargo', 'test', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--features', 'allocator-metrics', '--lib', '--bin', 'litchi-perf-baseline-alloc']),
    ('format', ['cargo', 'fmt', '--all', '--', '--check']),
    ('zip-format', ['cargo', 'fmt', '-p', 'soapberry-zip', '--', '--check']),
    ('harness-format', ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--', '--check']),
    ('zip-clippy', ['cargo', 'clippy', '--release', '--locked', '-p', 'soapberry-zip', '--all-targets', '--no-deps', '--', '-D', 'warnings']),
    ('zip-rustdoc', ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', '--locked', '-p', 'soapberry-zip', '--no-deps']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
    ('registry-strict', ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict']),
]

if __name__ == '__main__':
    selected = set(sys.argv[1:])
    for label, argv in COMMANDS:
        if selected and label not in selected:
            continue
        result = subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *argv], cwd=REPO, env=ENV)
        if result.returncode:
            raise SystemExit(result.returncode)
