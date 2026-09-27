"""Fresh serial ordinary harness build, bound to the recorded current source."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0772'


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def guard():
    source = json.loads((P / 'source.json').read_text())
    for files in (source['files'], json.loads((P / 'constraints.json').read_text()),
                  json.loads((P / 'workspace-inputs.json').read_text())):
        for name, digest in files.items():
            assert sha(ROOT / name) == digest, name
    assert subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip() == source['head']


if __name__ == '__main__':
    guard()
    assert not TARGET.exists() and not (P / 'build.json').exists()
    command = ['cargo', 'build', '--release', '--offline', '--locked',
               '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--bin', 'litchi-perf-baseline']
    env = os.environ | {'CARGO_TARGET_DIR': str(TARGET), 'CARGO_BUILD_JOBS': '2'}
    start = time.time()
    with (P / 'build-native.log').open('w') as log:
        result = subprocess.run(command, cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
    receipt = {'command': command, 'target': str(TARGET), 'started': start,
               'ended': time.time(), 'exit': result.returncode,
               'log_sha256': sha(P / 'build-native.log'), 'source_sha256': sha(P / 'source.json')}
    if result.returncode == 0:
        binary = TARGET / 'release/litchi-perf-baseline'
        receipt |= {'binary': str(binary), 'binary_sha256': sha(binary), 'binary_bytes': binary.stat().st_size}
    (P / 'build.json').write_text(json.dumps(receipt, indent=2) + '\n')
    assert result.returncode == 0, receipt
    guard()
    print('PASS current-source native harness build', flush=True)
