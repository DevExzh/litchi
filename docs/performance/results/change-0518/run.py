"""Build the unchanged DOCX example and bind its executable to current sources."""
import datetime
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path('/tmp/litchi-goal-0518')
TARGET = Path('/home/zhuhe/litchi-goal-0518-target')


def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2)
        stream.write('\n')


def sources():
    names = subprocess.check_output([
        'git', 'ls-files', '--cached', '--others', '--exclude-standard', '-z',
        'crates', 'tools/perf-baseline', 'Cargo.toml', 'Cargo.lock',
        'rust-toolchain.toml', '.cargo'], cwd=REPO).decode().split('\0')
    return {name: sha(REPO / name) for name in sorted(names)
            if name and Path(name).suffix in {'.rs', '.toml', '.lock'}}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--variant', choices=['baseline', 'candidate'], default='baseline')
    variant = parser.parse_args().variant
    directory = HERE / variant
    directory.mkdir()
    manifest = sources()
    write(directory / 'source-manifest.json', manifest)
    command = ['cargo', 'build', '--release', '--locked', '-p', 'litchi-docx',
               '--example', 'managed_paragraph_batch_perf']
    settings = dict(CARGO_TARGET_DIR=str(TARGET), CARGO_BUILD_JOBS='2',
                    CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='0',
                    CARGO_PROFILE_DEV_DEBUG='0', CARGO_PROFILE_TEST_DEBUG='0',
                    TMPDIR=str(SCRATCH))
    environment = dict(os.environ, **settings)
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with (directory / 'build.log').open('x') as stream:
        result = subprocess.run(['/usr/bin/time', '-v', *command], cwd=REPO,
                                env=environment, stdout=stream, stderr=subprocess.STDOUT)
    receipt = {'command': command, 'started_utc': start,
               'elapsed_seconds': time.monotonic() - tick,
               'exit_code': result.returncode, 'environment': settings,
               'source_unchanged': sources() == manifest,
               'source_manifest_sha256': sha(directory / 'source-manifest.json'),
               'log_sha256': sha(directory / 'build.log')}
    if result.returncode == 0:
        binary = SCRATCH / ('managed-paragraph-' + variant)
        assert not binary.exists()
        shutil.copy2(TARGET / 'release/examples/managed_paragraph_batch_perf', binary)
        receipt['binary'] = str(binary)
        receipt['binary_sha256'] = sha(binary)
    write(directory / 'build-receipt.json', receipt)
    print('DOCX', variant, 'build', result.returncode, flush=True)
    assert result.returncode == 0 and receipt['source_unchanged']


if __name__ == '__main__':
    main()
