#!/usr/bin/env python3
"""Freeze and execute separate total/phase DOCX append measurements."""
import subprocess
import sys
from common import ROOT, REPO, TEMP, ENV, ENV_KEYS, sha, meta, now, read, write


def plan():
    forward = [(instrumentation, count, mode)
               for instrumentation in ('normal', 'allocator')
               for count in (64, 8192, 131072)
               for mode in ('total', 'phases')]
    captures = []
    for repeat, sequence in ((1, forward), (2, list(reversed(forward)))):
        for instrumentation, count, mode in sequence:
            label = f'r{repeat}-{instrumentation}-{count}-{mode}'
            output = ROOT / 'captures' / label
            argv = ['/usr/bin/time', '-v', '-o', str(output.with_suffix('.resource')),
                    '/usr/bin/taskset', '-c', '2',
                    str(TEMP / instrumentation / 'docx_plain_paragraph_tail_append'),
                    '--mode', mode, '--counts', str(count), '--samples', '30',
                    '--warmups', '3', '--json', str(output.with_suffix('.report.json'))]
            captures.append(dict(label=label, instrumentation=instrumentation,
                                 count=count, mode=mode, repeat=repeat, argv=argv))
    return captures


def main():
    protocol_path = ROOT / 'protocol.json'
    if sys.argv[1:] == ['--freeze']:
        write(protocol_path, dict(
            schema='docx-plain-paragraph-tail-append-capture-v1',
            samples=30, warmups=3, cpu=2, append_count=1,
            comparison='baseline existing one-paragraph logical tail copy; total and phase observer modes are separate measurements',
            scope='source adapter and package construction, snapshot, edit staging, commit, sequential publication and owner drop; pre-existing corpus/oracle bytes excluded',
            normal_and_allocator_timings_separate=True,
            process_rss_includes_setup_oracles_and_teardown=True,
            phase_peaks_are_not_total_peaks=True,
            regression_review_percent=5, performance_claim='none',
            scripts={name: sha(ROOT / name) for name in ('common.py', 'capture.py')},
            environment={key: ENV[key] for key in ENV_KEYS}, captures=plan()))
        print('Frozen 24 captures / 720 operations.')
        return
    protocol = read(protocol_path)
    assert protocol['captures'] == plan()
    assert protocol['scripts'] == {name: sha(ROOT / name) for name in protocol['scripts']}
    binaries = read(ROOT / 'binaries.json')
    for spec in binaries.values():
        assert meta(spec['path']) == {key: spec[key] for key in ('bytes', 'sha256')}
    (ROOT / 'captures').mkdir(exist_ok=True)
    selected = set(sys.argv[1:])
    assert selected <= {capture['label'] for capture in protocol['captures']}
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
