#!/usr/bin/env python3
"""Run applicable benchmark and unchanged DOCX contract gates serially."""
import subprocess
import sys
from common import ROOT, REPO, ENV

MANIFEST = ['--manifest-path', 'tools/perf-baseline/Cargo.toml']
TARGETS = ['--lib', '--bin', 'docx_plain_paragraph_tail_append']
GATES = {
    'harness-tests-final': ['cargo', 'test', '--release', '--locked', *MANIFEST,
                      '--features', 'allocator-metrics', *TARGETS],
    'harness-clippy-final': ['cargo', 'clippy', '--release', '--locked', *MANIFEST,
                       '--features', 'allocator-metrics', *TARGETS,
                       '--no-deps', '--', '-D', 'warnings'],
    'harness-format-final': ['cargo', 'fmt', *MANIFEST, '--', '--check'],
    'harness-rustdoc': ['env', 'RUSTDOCFLAGS=-D warnings', 'cargo', 'doc',
                        '--locked', *MANIFEST, '--lib', '--no-deps'],
    'docx-paragraph-copy': ['cargo', 'test', '--release', '--locked', '-p',
                            'litchi-docx', '--test', 'source_backed_paragraph_copy'],
    'docx-paragraph-removal': ['cargo', 'test', '--release', '--locked', '-p',
                               'litchi-docx', '--test', 'source_backed_paragraph_removal'],
    'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
    'registry-strict': ['python3', '-B', 'tools/check_perf_claims.py', '--registry',
                         'docs/performance/claim-registry-v1.json', '--repo-root',
                         '.', '--evidence-root', '.', '--mode', 'strict'],
    'evidence-tests-final': ['python3', '-B', '-m', 'unittest', 'discover', '-s',
                       'docs/performance/results/change-0479', '-p', 'test_evidence.py'],
    'analyze-final': ['python3', '-B', 'docs/performance/results/change-0479/analyze.py'],
}


if __name__ == '__main__':
    for label in sys.argv[1:] or list(GATES):
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label,
                        *GATES[label]], cwd=REPO, env=ENV, check=True)
