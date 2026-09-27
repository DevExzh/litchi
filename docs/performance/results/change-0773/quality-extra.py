"""Broader integration regression gates, serial and source-bound.

Continue after a failed test command to retain the full regression picture;
any failure keeps overall status failed and needs an explicit disposition.
"""
import json
import os
import subprocess
import time

from quality import P, ROOT, TARGET, OWNERS, census, sha


if __name__ == '__main__':
    before = census()
    attempt = 0
    while (P / f'broader-{attempt}').exists():
        attempt += 1
    folder = P / f'broader-{attempt}'
    folder.mkdir()
    (folder / 'source.json').write_text(json.dumps(before, indent=2) + '\n')
    formats = [arg for name in OWNERS[3:] for arg in ('-p', name)]
    dependents = [arg for name in ('litchi-ole-common', 'litchi-ooxml-common', 'litchi-crypto',
                                  'litchi-vba', 'litchi-sign') for arg in ('-p', name)]
    locked = ['--offline', '--locked']
    commands = [
        ['cargo', 'test', *formats, *locked, '--all-features', '--', '--test-threads=2'],
        ['cargo', 'test', *dependents, *locked, '--all-features', '--', '--test-threads=2'],
        ['cargo', 'test', '-p', 'litchi', *locked, '--features',
         'doc,docx,ppt,pptx,xls,xlsx,xlsb,odt', '--', '--test-threads=2'],
        ['cargo', 'test', '--manifest-path', 'tools/perf-baseline/Cargo.toml', *locked,
         '--lib', '--', '--test-threads=2'],
    ]
    env = os.environ | {'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2',
                       'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0',
                       'RUSTDOCFLAGS': '-D warnings', 'PYTHONDONTWRITEBYTECODE': '1'}
    rows = []
    for index, command in enumerate(commands):
        log = folder / f'{index:02d}.log'
        start = time.time()
        with log.open('w') as stream:
            result = subprocess.run(command, cwd=ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT)
        rows.append({'command': command, 'exit': result.returncode, 'started': start,
                     'ended': time.time(), 'log': str(log.relative_to(P)), 'sha256': sha(log)})
        receipt = {'source_file': str((folder / 'source.json').relative_to(P)),
                   'source_sha256': sha(folder / 'source.json'), 'rows': rows}
        (folder / 'manifest.json').write_text(json.dumps(receipt, indent=2) + '\n')
        assert census() == before, 'source changed during broader gates'
        print(f'broader gate {index + 1}/{len(commands)} exit {result.returncode}', flush=True)
    receipt['status'] = 'pass' if all(row['exit'] == 0 for row in rows) else 'failed'
    (P / 'broader.json').write_text(json.dumps(receipt, indent=2) + '\n')
    raise SystemExit(0 if receipt['status'] == 'pass' else 1)
