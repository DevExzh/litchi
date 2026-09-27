"""Retest the exporter after the two-line Clippy correction; retain lineage."""
import os
from pathlib import Path
import subprocess
import time
import custody as c

if __name__ == '__main__':
    p = c.P
    source = c.read(p / 'source.json')
    old = c.read(p / 'quality-0/source.json')
    assert c.census() == source
    changed = [name for name in old['files'] if old['files'][name] != source['files'][name]]
    assert changed == ['tools/perf-baseline/src/ordinary_save.rs']
    before = subprocess.check_output(['git', 'show', old['revision'] + ':' + changed[0]], cwd=c.ROOT).decode()
    after = (c.ROOT / changed[0]).read_text()
    assert before.count('.then(|| corpus.pptx_target)') == 2
    assert before.replace('.then(|| corpus.pptx_target)', '.then_some(corpus.pptx_target)') == after
    diff = p / 'library-followup.diff'
    assert not diff.exists()
    diff.write_bytes(subprocess.check_output(['git', 'diff', old['revision'], source['revision'], '--', changed[0]], cwd=c.ROOT))
    command = ['cargo', 'test', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--offline', '--locked', '--features', 'allocator-metrics', '--lib', 'ordinary_save::tests::artifact_export_generated_matrix_matches_reference_and_refuses_reuse', '--', '--exact', '--test-threads=1']
    env = os.environ | {'CARGO_TARGET_DIR': c.read(p / 'plan.json')['target'], 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0'}
    started = time.time()
    log = p / 'library-followup.log'
    with log.open('x') as f:
        run = subprocess.run(command, cwd=c.ROOT, env=env, stdout=f, stderr=subprocess.STDOUT)
    c.write(p / 'library-followup.json', {'command': command, 'exit_code': run.returncode,
            'started': started, 'ended': time.time(), 'log': c.artifact(log),
            'source_sha256': c.sha(p / 'source.json'), 'predecessor_source_sha256': c.sha(p / 'quality-0/source.json'),
            'only_source_difference': c.artifact(diff), 'replacement_count': 2,
            'scope': 'Final-source focused exporter retest after exact two-line then_some Clippy correction; full library suite passed on predecessor.'})
    assert run.returncode == 0 and c.census() == source
