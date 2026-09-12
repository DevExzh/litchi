"""Serial correctness and policy gates; retain each exact command and exit code."""
import datetime
import argparse
import os
import subprocess
import time
from run import HERE, REPO, TARGET, SCRATCH, sha, sources, write

COMMANDS = [
    ('fmt', ['cargo', 'fmt', '--all', '--check']),
    ('ooxml-tests', ['cargo', 'test', '--locked', '-p', 'litchi-opc', '-p', 'litchi-docx', '-p', 'litchi-xlsx', '-p', 'litchi-pptx', '-p', 'litchi-xlsb', '--all-features', '--', '--test-threads=2']),
    ('workspace-check', ['cargo', 'check', '--locked', '--workspace', '--all-features']),
    ('clippy', ['cargo', 'clippy', '--locked', '-p', 'litchi-opc', '-p', 'litchi-docx', '--all-features', '--lib', '--', '-D', 'warnings']),
    ('rustdoc', ['cargo', 'doc', '--locked', '-p', 'litchi-opc', '-p', 'litchi-docx', '--all-features', '--no-deps']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
    ('zip64-generate', ['python3', '-B', 'docs/performance/results/change-0415/interop.py', 'generate', str(SCRATCH / 'independent-zip64.zip')]),
    ('zip64-test', ['cargo', 'test', '--locked', '-p', 'litchi-opc', '--all-features', '--test', 'external_zip64_source', '--', '--ignored', '--nocapture']),
    ('claims', ['python3', '-B', 'tools/check_perf_claims.py', '--registry', 'docs/performance/claim-registry-v1.json', '--repo-root', '.', '--evidence-root', '.', '--mode', 'strict']),
]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--only', choices=[name for name, _ in COMMANDS])
    parser.add_argument('--directory', default='checks')
    args = parser.parse_args()
    directory = HERE / args.directory
    directory.mkdir(exist_ok=True)
    manifest = sources()
    assert manifest == __import__('json').loads((HERE / 'candidate/source-manifest.json').read_text())
    settings = dict(CARGO_TARGET_DIR=str(TARGET), CARGO_BUILD_JOBS='2', CARGO_INCREMENTAL='0', CARGO_PROFILE_RELEASE_DEBUG='0', CARGO_PROFILE_DEV_DEBUG='0', CARGO_PROFILE_TEST_DEBUG='0', TMPDIR=str(SCRATCH), RUSTDOCFLAGS='-D warnings')
    settings['LITCHI_0415_PYTHON_ZIP'] = str(SCRATCH / 'independent-zip64.zip')
    for name, command in COMMANDS:
        if args.only and args.only != name:
            continue
        start = datetime.datetime.now(datetime.timezone.utc).isoformat()
        tick = time.monotonic()
        log = directory / (name + '.log')
        with log.open('x') as stream:
            result = subprocess.run(command, cwd=REPO, env=dict(os.environ, **settings), stdout=stream, stderr=subprocess.STDOUT)
        unchanged = sources() == manifest
        write(directory / (name + '.json'), dict(command=command, started_utc=start, elapsed_seconds=time.monotonic()-tick, exit_code=result.returncode, source_unchanged=unchanged, source_manifest_sha256=sha(HERE/'candidate/source-manifest.json'), log_sha256=sha(log), environment=settings))
        print(name, result.returncode, flush=True)
        assert unchanged
        if result.returncode:
            print(log.read_text()[-6000:], flush=True)
            raise SystemExit(result.returncode)

if __name__ == '__main__':
    main()
