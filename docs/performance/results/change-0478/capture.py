#!/usr/bin/env python3
"""Capture isolated, matched storage policies under the coordinator's CPU lock."""
import json
from pathlib import Path
import subprocess
import sys

from common import ROOT, REPO, TEMP, ENV, ENV_KEYS, sha, meta, now, read, write


def plan():
    forward = [(instrumentation, count, policy)
               for instrumentation in ('normal', 'allocator')
               for count in (8, 256, 8192)
               for policy in ('control', 'spool')]
    captures = []
    for repeat, sequence in ((1, forward), (2, list(reversed(forward)))):
        for instrumentation, count, policy in sequence:
            label = f'r{repeat}-{instrumentation}-{count}-{policy}'
            output = ROOT / 'captures' / label
            binary = TEMP / instrumentation / 'pptx_metadata_spool'
            argv = ['/usr/bin/time', '-v', '-o', str(output.with_suffix('.resource')),
                    '/usr/bin/taskset', '-c', '2', str(binary),
                    '--mode', policy, '--counts', str(count),
                    '--samples', '30', '--warmups', '3', '--repeats', '1',
                    '--max-spool-bytes', str(64 * 1024 * 1024),
                    '--spool-buffer-bytes', '16384',
                    '--json', str(output.with_suffix('.report.json')),
                    '--spool-dir', str(TEMP / 'spools' / label)]
            captures.append(dict(label=label, instrumentation=instrumentation,
                                 count=count, policy=policy,
                                 repeat=repeat, argv=argv))
    return captures


def main():
    protocol_path = ROOT / 'protocol.json'
    if sys.argv[1:] == ['--freeze']:
        write(protocol_path, dict(
            schema='pptx-metadata-spool-capture-v1', samples=30, warmups=3,
            preliminary=dict(path='preliminary/manifest.json', **meta(ROOT / 'preliminary/manifest.json')),
            comparison='same-source public PPTX default writer versus generated-plan explicit-file-spool writer',
            preceding_evidence='change-0476/summary.json', cpu=2,
            preceding_evidence_sha256=sha(ROOT.parent / 'change-0476' / 'summary.json'),
            scope='complete public PPTX writer and plan creation, generated text, slides, finalization, file open/close; '
                  'input corpus preparation, output digest finalization, oracle and cleanup excluded',
            normal_and_allocator_timings_separate=True,
            process_rss_includes_setup_oracles_and_teardown=True,
            regression_review_percent=5,
            memory_gate='Across 8/256/8192 slides, generated-route operation peak range must be within 1% of its smallest-count peak, with zero failed allocations and zero live exit delta; larger growth requires source/heap investigation before any bounded-window claim. Descriptor, active-name and integer-width effects must be accounted for.',
            scripts={name: sha(ROOT / name) for name in ('common.py', 'capture.py')},
            environment={key: ENV[key] for key in ENV_KEYS}, captures=plan()))
        print('frozen 24 captures', flush=True)
        return
    protocol = read(protocol_path)
    assert protocol['captures'] == plan()
    assert protocol['scripts'] == {name: sha(ROOT / name) for name in protocol['scripts']}
    binaries = read(ROOT / 'binaries.json')
    for spec in binaries.values():
        assert meta(spec['path']) == {key: spec[key] for key in ('bytes', 'sha256')}
    (ROOT / 'captures').mkdir(exist_ok=True)
    selected = set(sys.argv[1:])
    for capture in protocol['captures']:
        label = capture['label']
        if selected and label not in selected:
            continue
        output = ROOT / 'captures' / label
        started = dict(capture=capture, cwd=str(REPO), started_utc=now(),
                       protocol_sha256=sha(protocol_path),
                       binary=binaries[capture['instrumentation']],
                       environment={key: ENV[key] for key in ENV_KEYS})
        write(output.with_suffix('.started.json'), started)
        with output.with_suffix('.stdout').open('xb') as out, output.with_suffix('.stderr').open('xb') as err:
            result = subprocess.run(capture['argv'], cwd=REPO, env=ENV, stdout=out, stderr=err)
        record = dict(started, exit_code=result.returncode, finished_utc=now(),
                      artifacts={output.with_suffix(suffix).name: meta(output.with_suffix(suffix))
                                 for suffix in ('.stdout', '.stderr', '.resource', '.report.json')
                                 if output.with_suffix(suffix).is_file()})
        write(output.with_suffix('.json'), record)
        print(label, result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
