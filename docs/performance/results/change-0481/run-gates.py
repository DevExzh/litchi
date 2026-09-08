#!/usr/bin/env python3
"""Run applicable candidate validation serially under the batch CPU lock."""
import subprocess
import sys
from common import ROOT, REPO, ENV
GATES = {
    'format-docx': ['cargo','fmt','--package','litchi-docx','--','--check'],
    'docx-tests': ['cargo','test','--release','--locked','-p','litchi-docx','--lib','--test','source_backed_paragraph_copy','--test','source_backed_paragraph_removal'],
    'opc-shared-tests': ['cargo','test','--release','--locked','-p','litchi-opc','--lib','shared_'],
    'harness-tests': ['cargo','test','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--lib','docx_plain_paragraph_tail_append::tests'],
    'clippy-all-features': ['cargo','clippy','--release','--locked','-p','litchi-docx','--all-features','--lib','--tests','--no-deps','--','-D','warnings'],
    'rustdoc': ['env','RUSTDOCFLAGS=-D warnings','cargo','doc','--locked','-p','litchi-docx','--lib','--no-deps'],
    'boundaries': ['python3','-B','tools/check_crate_boundaries.py'],
    'registry-strict': ['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--evidence-root','.','--mode','strict'],
    'evidence-tests-final': ['python3','-B','-m','unittest','discover','-s','docs/performance/results/change-0481','-p','test_evidence.py'],
    'analyze-final': ['python3','-B','docs/performance/results/change-0481/analyze.py'],
}
if __name__ == '__main__':
    for label in sys.argv[1:]:
        subprocess.run([sys.executable,'-B',str(ROOT/'gate.py'),label,*GATES[label]],cwd=REPO,env=ENV,check=True)
