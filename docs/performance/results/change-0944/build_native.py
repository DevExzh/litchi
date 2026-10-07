#!/usr/bin/env python3
import hashlib
import json
import os
import pathlib
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone

PACKET = pathlib.Path('/home/zhuhe/litchi-performance-0944')
SOURCE = PACKET / 'source-build'
TARGET = PACKET / 'target-native'
TMP = PACKET / 'tmp' / 'build-native'
BUILD = PACKET / 'build'
BUILD.mkdir(parents=True, exist_ok=True)
TMP.mkdir(parents=True, exist_ok=True)
TARGET.mkdir(parents=True, exist_ok=True)
START_PATH = BUILD / 'native-build-start.json'
TERM_PATH = BUILD / 'native-build-terminal.json'
STDOUT = BUILD / 'native-build.stdout.log'
STDERR = BUILD / 'native-build.stderr.log'
BINARY = TARGET / 'release' / 'litchi-perf-baseline'
argv = [
    '/home/zhuhe/.cargo/bin/cargo', 'build', '--release', '--locked', '--offline',
    '--manifest-path', str(SOURCE / 'tools/perf-baseline/Cargo.toml'),
    '--bin', 'litchi-perf-baseline', '-j', '2',
]
env = {
    'PATH': '/home/zhuhe/.cargo/bin:/usr/bin:/bin',
    'HOME': '/home/zhuhe',
    'CARGO_HOME': '/home/zhuhe/.cargo',
    'RUSTUP_HOME': '/home/zhuhe/.rustup',
    'CARGO_TARGET_DIR': str(TARGET),
    'CARGO_NET_OFFLINE': 'true',
    'CARGO_INCREMENTAL': '0',
    'CARGO_BUILD_JOBS': '2',
    'CARGO_PROFILE_RELEASE_CODEGEN_UNITS': '1',
    'CARGO_PROFILE_RELEASE_DEBUG': '1',
    'CARGO_PROFILE_RELEASE_LTO': 'fat',
    'CARGO_PROFILE_RELEASE_OPT_LEVEL': '3',
    'RUSTUP_TOOLCHAIN': '1.95.0',
    'LANG': 'C', 'LC_ALL': 'C', 'TZ': 'UTC', 'RUST_BACKTRACE': '0',
    'TMPDIR': str(TMP), 'TEMP': str(TMP), 'TMP': str(TMP),
}
missing = [str(SOURCE / 'tools/perf-baseline/Cargo.lock'), str(SOURCE / '.cargo/config.toml')]
for item in missing:
    if not pathlib.Path(item).is_file():
        raise SystemExit(f'missing bound build input: {item}')
for name in ('RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'CARGO_BUILD_RUSTC_WRAPPER', 'RUSTC_WRAPPER'):
    if name in env:
        raise SystemExit(f'forbidden custom flag unexpectedly set: {name}')

def now(): return datetime.now(timezone.utc).isoformat()
def digest(path):
    h=hashlib.sha256(); n=0
    with path.open('rb') as f:
        while True:
            b=f.read(1024*1024)
            if not b: break
            h.update(b); n += len(b)
    return {'path':str(path.relative_to(PACKET)),'bytes':n,'sha256':h.hexdigest()}
def write_json(path, obj):
    tmp=path.with_suffix(path.suffix+'.tmp')
    tmp.write_text(json.dumps(obj, indent=2, sort_keys=True)+'\n')
    with tmp.open('rb') as f: os.fsync(f.fileno())
    os.replace(tmp,path)
    d=os.open(path.parent, os.O_DIRECTORY); os.fsync(d); os.close(d)

runner = pathlib.Path(__file__).resolve()
start = now()
write_json(START_PATH, {
    'schema':'litchi-xlsx-profile-build-start-v1',
    'started_utc':start,
    'runner':str(runner.relative_to(PACKET)),
    'runner_sha256':hashlib.sha256(runner.read_bytes()).hexdigest(),
    'cwd':str(SOURCE), 'cwd_relative':'source-build', 'argv':argv,
    'env':env, 'unset_env':['RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS','RUSTC_WRAPPER','CARGO_BUILD_RUSTC_WRAPPER'],
    'deadline_seconds':900, 'process_group':True, 'start_new_session':True,
    'inputs':{
      'source_manifest':'source-build-manifest.json',
      'source_manifest_sha256':hashlib.sha256((PACKET/'source-build-manifest.json').read_bytes()).hexdigest(),
      'tracked_lock':'source-build/tools/perf-baseline/Cargo.lock',
      'tracked_lock_sha256':hashlib.sha256((SOURCE/'tools/perf-baseline/Cargo.lock').read_bytes()).hexdigest(),
      'root_manifest':'source-build/tools/perf-baseline/Cargo.toml',
      'root_manifest_sha256':hashlib.sha256((SOURCE/'tools/perf-baseline/Cargo.toml').read_bytes()).hexdigest(),
      'cargo_config':'source-build/.cargo/config.toml',
      'cargo_config_sha256':hashlib.sha256((SOURCE/'.cargo/config.toml').read_bytes()).hexdigest(),
    },
    'stdout':str(STDOUT.relative_to(PACKET)), 'stderr':str(STDERR.relative_to(PACKET)),
    'terminal_receipt':str(TERM_PATH.relative_to(PACKET)),
})
with STDOUT.open('wb') as out, STDERR.open('wb') as err:
    proc = subprocess.Popen(argv, cwd=SOURCE, env=env, stdout=out, stderr=err, start_new_session=True)
    pid=proc.pid
    start_obj=json.loads(START_PATH.read_text()); start_obj['pid']=pid; start_obj['process_group_id']=os.getpgid(pid); write_json(START_PATH,start_obj)
    deadline=time.monotonic()+900
    timed_out=False
    while True:
        rc=proc.poll()
        if rc is not None: break
        if time.monotonic() >= deadline:
            timed_out=True
            try: os.killpg(os.getpgid(pid), signal.SIGTERM)
            except ProcessLookupError: pass
            try: proc.wait(timeout=10)
            except subprocess.TimeoutExpired:
                try: os.killpg(os.getpgid(pid), signal.SIGKILL)
                except ProcessLookupError: pass
                proc.wait()
            rc=proc.returncode
            break
        time.sleep(0.25)
end=now()
term={
    'schema':'litchi-xlsx-profile-build-terminal-v1','started_utc':start,'terminal_utc':end,
    'pid':pid,'process_group_id':start_obj['process_group_id'],'returncode':rc,'timed_out':timed_out,
    'deadline_seconds':900,'stdout':digest(STDOUT),'stderr':digest(STDERR),
    'binary':digest(BINARY) if BINARY.is_file() else None,
    'terminal_before_parse':True,
}
write_json(TERM_PATH,term)
print(json.dumps(term,sort_keys=True))
if rc != 0 or timed_out or not BINARY.is_file(): sys.exit(1)
