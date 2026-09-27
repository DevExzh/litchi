"""Root-only untimed export with build/source/fixture custody."""
from pathlib import Path
import subprocess
import time
import custody as c


def main():
    p = c.P
    plan = c.read(p / 'plan.json')
    source = c.read(p / 'source.json')
    build = c.read(p / 'build.json')
    binary = build['binaries']['export']
    assert build['source_sha256'] == c.sha(p / 'source.json')
    fixtures = {row['path']: {'bytes': row['bytes'], 'sha256': row['sha256']}
                for row in plan['corpora'] if row['path']}

    def guard():
        assert c.census() == source
        assert c.artifact(Path(binary['path'])) == binary
        for name, expected in fixtures.items():
            path = c.ROOT / name
            assert path.stat().st_size == expected['bytes'] and c.sha(path) == expected['sha256']

    guard()
    out = p / 'artifacts-0'
    receipt = p / 'export.json'
    assert not out.exists() and not receipt.exists()
    command = [binary['path'], '--output', str(out), '--filesystem-root', plan['filesystem_root']]
    for name in fixtures:
        command.extend(['--ooxml-file', str(c.ROOT / name)])
    started = time.time()
    with (p / 'export.stdout').open('x') as stdout, (p / 'export.stderr').open('x') as stderr:
        run = subprocess.run(command, cwd=c.ROOT, stdout=stdout, stderr=stderr)
    guard()
    record = {'command': command, 'exit_code': run.returncode, 'started': started,
              'ended': time.time(), 'binary': binary, 'source_sha256': c.sha(p / 'source.json'),
              'plan_sha256': c.sha(p / 'plan.json'), 'build_sha256': c.sha(p / 'build.json'),
              'runner_sha256': c.sha(Path(__file__)), 'fixtures': fixtures,
              'stdout': c.artifact(p / 'export.stdout'), 'stderr': c.artifact(p / 'export.stderr')}
    if run.returncode == 0:
        record['manifest'] = c.artifact(out / 'manifest.json')
    c.write(receipt, record)
    assert run.returncode == 0, record


if __name__ == '__main__':
    main()
