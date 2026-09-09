#!/usr/bin/env python3
"""Serial, retained checks for the bounded XML audit reader."""
import subprocess
import sys
from common import ROOT, REPO, ENV

GATES = {
    'format-xml': ['cargo', 'fmt', '--package', 'xml-minifier', '--', '--check'],
    'xml-tests': ['cargo', 'test', '--release', '--locked', '-p', 'xml-minifier'],
    'xml-no-default': ['cargo', 'check', '--locked', '-p', 'xml-minifier', '--no-default-features'],
    'xml-clippy': ['cargo', 'clippy', '--release', '--locked', '-p', 'xml-minifier', '--all-features', '--all-targets', '--no-deps', '--', '-D', 'warnings'],
    'xml-rustdoc': ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', '--locked', '-p', 'xml-minifier', '--no-deps'],
    'format-opc': ['cargo', 'fmt', '--package', 'litchi-opc', '--', '--check'],
    'opc-tests': ['cargo', 'test', '--release', '--locked', '-p', 'litchi-opc', '--all-features'],
    'opc-clippy': ['cargo', 'clippy', '--release', '--locked', '-p', 'litchi-opc', '--all-features', '--all-targets', '--no-deps', '--', '-D', 'warnings'],
    'opc-rustdoc': ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', '--locked', '-p', 'litchi-opc', '--all-features', '--no-deps'],
    'format-zip': ['cargo', 'fmt', '--package', 'soapberry-zip', '--', '--check'],
    'zip-tests': ['cargo', 'test', '--release', '--locked', '-p', 'soapberry-zip', '--all-features'],
    'zip-clippy': ['cargo', 'clippy', '--release', '--locked', '-p', 'soapberry-zip', '--all-features', '--all-targets', '--no-deps', '--', '-D', 'warnings'],
    'zip-rustdoc': ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc', '--locked', '-p', 'soapberry-zip', '--no-deps'],
    'workspace-check': ['cargo', 'check', '--locked', '--workspace'],
    'harness-tests': ['cargo', 'test', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--lib', 'xml_stream_audit'],
    'harness-clippy': ['cargo', 'clippy', '--release', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--bin', 'xml_stream_audit', '--features', 'allocator-metrics', '--no-deps', '--', '-D', 'warnings'],
    'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
    'registry-strict': ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict'],
    'evidence-tests-final': ['python3', '-B', 'docs/performance/results/change-0482/test_evidence.py'],
    'evidence-tests-final-v2': ['python3', '-B', 'docs/performance/results/change-0482/test_evidence.py'],
    'analyze-final': ['python3', '-B', 'docs/performance/results/change-0482/analyze.py'],
    'xml-fuzz-build': ['python3', '-B', 'docs/performance/results/change-0482/fuzz.py', 'build'],
    'xml-fuzz-smoke': ['python3', '-B', 'docs/performance/results/change-0482/fuzz.py', 'smoke'],
}

# The first XML/OPC/ZIP gate attempts are retained as developmental history.
# Fresh labels make the post-rustdoc-fix results distinguishable and required.
GATES.update({
    'format-xml-final': GATES['format-xml'],
    'xml-tests-final': GATES['xml-tests'],
    'xml-no-default-final': GATES['xml-no-default'],
    'xml-clippy-final': GATES['xml-clippy'],
    'xml-rustdoc-final': GATES['xml-rustdoc'],
    'format-opc-final': GATES['format-opc'],
    'format-zip-final': GATES['format-zip'],
    'zip-tests-final': GATES['zip-tests'],
    'zip-clippy-final': GATES['zip-clippy'],
    'zip-rustdoc-final': GATES['zip-rustdoc'],
})

if __name__ == '__main__':
    for label in sys.argv[1:]:
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *GATES[label]],
                       cwd=REPO, env=ENV, check=True)
