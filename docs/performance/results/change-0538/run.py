"""Serial correctness-only harness checks and allocator schema smoke captures."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
TARGET = Path('/home/zhuhe/litchi-goal-0538-target')
ENV = dict(os.environ, CARGO_TARGET_DIR=str(TARGET), CARGO_BUILD_JOBS='4',
           CARGO_INCREMENTAL='0', CARGO_PROFILE_DEV_DEBUG='0', CARGO_PROFILE_TEST_DEBUG='0',
           TMPDIR=str(TARGET/'tmp'))
MANIFEST = ['--manifest-path', 'tools/perf-baseline/Cargo.toml', '--locked']

def run(name, command, extra=None):
    assert not (HERE/'SHA256SUMS').exists(), 'sealed evidence is read-only'
    receipt = HERE/(name+'.receipt.json')
    assert not receipt.exists(), name
    (TARGET/'tmp').mkdir(parents=True, exist_ok=True)
    started = time.time()
    with (HERE/(name+'.log')).open('w') as log:
        process = subprocess.Popen(command, cwd=REPO, env=dict(ENV, **(extra or {})),
                                   stdout=log, stderr=subprocess.STDOUT)
        code = process.wait()
    receipt.write_text(json.dumps(dict(command=command, started=started,
        finished=time.time(), pid=process.pid, exit_code=code,
        environment={k: v for k, v in dict(ENV, **(extra or {})).items()
                     if k.startswith('CARGO_') or k in ['TMPDIR','RUSTDOCFLAGS']},
        source_sha256=hashlib.sha256((REPO/'tools/perf-baseline/src/lib.rs').read_bytes()).hexdigest()),
        indent=2, sort_keys=True)+'\n')
    print(name, code, flush=True)
    assert code == 0, name

def quality():
    run('test-xlsx', ['cargo','test',*MANIFEST,'--all-features','--lib','xlsx_', '--','--test-threads=1'])
    remaining_quality()

def resume_quality():
    run('test-xlsx-tmpdir', ['cargo','test',*MANIFEST,'--all-features','--lib',
        'filesystem::tests::padded_xlsx_archive_keeps_typed_semantics_and_reports_aligned_hash',
        '--','--exact','--test-threads=1'])
    remaining_quality()

def remaining_quality():
    run('test-allocator', ['cargo','test',*MANIFEST,'--all-features','--lib','allocation_metrics::tests', '--','--test-threads=1'])
    run('test-wrapper', ['cargo','test',*MANIFEST,'--all-features','--bin','litchi-perf-baseline-alloc'])
    run('check', ['cargo','check',*MANIFEST,'--all-features','--all-targets'])
    run('clippy', ['cargo','clippy',*MANIFEST,'--all-features','--lib','--bins','--no-deps','--','-D','warnings'])
    run('rustdoc', ['cargo','doc',*MANIFEST,'--all-features','--no-deps'], {'RUSTDOCFLAGS':'-D warnings'})
    run('fmt', ['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml','--all','--','--check'])
    run('boundaries', ['python3','-B','tools/check_crate_boundaries.py'])
    run('build-smoke', ['cargo','build',*MANIFEST,'--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc'])

def capture():
    cases = ','.join('xlsx_source_backed_'+managed+'cell_values_'+edit+'_edit_save'
                     for managed in ['', 'managed_']
                     for edit in ['one', 'one_percent', 'batch', 'multi_sheet'])
    for binary in ['litchi-perf-baseline', 'litchi-perf-baseline-alloc']:
        for shape in ['medium','dense-sparse']:
            name = ('alloc' if binary.endswith('-alloc') else 'normal')+'-'+shape
            run(name, [str(TARGET/'debug'/binary),'--case',cases,
                '--xlsx-cell-crud-shape',shape,'--warmup','1','--samples','3',
                '--json',str(HERE/(name+'.json'))])

def integration():
    run('test-integration', ['cargo','test',*MANIFEST,'--features','allocator-metrics',
                            '--test','xlsx_planning_allocations','--','--test-threads=1'])

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('phase', choices=['quality','resume_quality','integration','capture'])
    args = parser.parse_args()
    globals()[args.phase]()
