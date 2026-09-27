"""Root-only DOCX owner gates after the preservation correction."""
import os
from pathlib import Path
import subprocess
import time
import custody as c

if __name__ == '__main__':
    p = c.P
    source = c.census()
    i = 0
    while (p / f'docx-quality-{i}').exists():
        i += 1
    out = p / f'docx-quality-{i}'
    out.mkdir()
    c.write(out / 'source.json', source)
    common = ['-p', 'litchi-docx', '--offline', '--locked', '--all-features']
    commands = [
        ['cargo', 'fmt', '-p', 'litchi-docx', '--', '--check'],
        ['cargo', 'check', *common, '--all-targets'],
        ['cargo', 'test', *common, '--', '--test-threads=2'],
        ['cargo', 'clippy', *common, '--lib', '--', '-D', 'warnings'],
        ['cargo', 'doc', *common, '--no-deps'],
    ]
    env = os.environ | {'CARGO_TARGET_DIR': c.read(p / 'plan.json')['target'], 'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0', 'CARGO_PROFILE_DEV_DEBUG': '0', 'RUSTDOCFLAGS': '-D warnings'}
    rows = []
    for index, command in enumerate(commands):
        log = out / f'{index:02}.log'
        started = time.time()
        with log.open('x') as f:
            result = subprocess.run(command, cwd=c.ROOT, env=env, stdout=f, stderr=subprocess.STDOUT)
        rows.append({'command': command, 'exit_code': result.returncode, 'started': started, 'ended': time.time(), 'log': str(log.relative_to(p)), 'log_sha256': c.sha(log)})
        c.write(out / 'checks.json', rows)
        print(f'DOCX gate {index+1}/5 exit {result.returncode}', flush=True)
        assert result.returncode == 0 and c.census() == source
    c.write(p / 'docx-quality.json', {'source_sha256': c.sha(out / 'source.json'), 'source_revision': source['revision'], 'rows': rows})
