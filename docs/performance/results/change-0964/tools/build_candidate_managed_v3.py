#!/usr/bin/env python3
from __future__ import annotations
import hashlib
import json
import os
import shutil
import signal
import stat
import subprocess
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path('/home/zhuhe/litchi-performance-0964')
SOURCE = ROOT / 'source-candidate'
MANIFEST = ROOT / 'source-manifest-candidate-v4.json'
TARGET = ROOT / 'target-candidate-managed-v3'
BUILD = ROOT / 'builds' / 'candidate-managed-v3'
RECEIPTS = ROOT / 'receipts'
LOGS = ROOT / 'logs'
TMP = ROOT / 'tmp' / 'candidate-managed-v3'
EXECUTABLES = ROOT / 'executables'
DEADLINE_SECONDS = 1800
COMMAND = [
    '/home/zhuhe/.cargo/bin/cargo', '+1.95.0', 'build', '--release', '--locked', '--offline',
    '--manifest-path', str(SOURCE / 'tools/perf-baseline/Cargo.toml'),
    '--features', 'allocator-metrics', '--bin', 'xlsx_managed_compression_alloc',
    '--example', 'xlsx_managed_compression_benchmark',
]
ENV_OVERRIDES = {
    'PATH': '/home/zhuhe/.cargo/bin:/usr/bin:/bin',
    'LANG': 'C', 'LC_ALL': 'C', 'TZ': 'UTC', 'RUST_BACKTRACE': '0',
    'CARGO_TERM_COLOR': 'never', 'CARGO_NET_OFFLINE': 'true',
    'CARGO_BUILD_JOBS': '2', 'CARGO_INCREMENTAL': '0',
    'CARGO_TARGET_DIR': str(TARGET), 'RUSTUP_TOOLCHAIN': '1.95.0',
    'CARGO_PROFILE_RELEASE_OPT_LEVEL': '3',
    'CARGO_PROFILE_RELEASE_LTO': 'fat',
    'CARGO_PROFILE_RELEASE_CODEGEN_UNITS': '1',
    'CARGO_PROFILE_RELEASE_DEBUG': '1',
    'TMPDIR': str(TMP),
}
REMOVED_FLAGS = ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'RUSTC_WRAPPER', 'RUSTC_WORKSPACE_WRAPPER', 'RUSTC', 'RUSTDOC', 'CARGO_BUILD_RUSTC_WRAPPER', 'CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER', 'CARGO_BUILD_RUSTFLAGS']
BINARY_PATHS = {
    'xlsx_managed_compression_alloc': TARGET / 'release' / 'xlsx_managed_compression_alloc',
    'xlsx_managed_compression_benchmark': TARGET / 'release' / 'examples' / 'xlsx_managed_compression_benchmark',
}
RETAINED_PATHS = {
    'xlsx_managed_compression_alloc': EXECUTABLES / 'xlsx_managed_compression_alloc.v3',
    'xlsx_managed_compression_benchmark': EXECUTABLES / 'xlsx_managed_compression_benchmark.v3',
}
RECEIPT_NAMES = [
    'candidate-managed-v3-build-start.json',
    'candidate-managed-v3-build-terminal.json',
    'candidate-managed-v3-build-binaries.json',
    'candidate-managed-v3-build-result.json',
    'candidate-managed-v3-retention.json',
]

def utc() -> str:
    return datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')

def digest(path: Path) -> dict:
    h = hashlib.sha256()
    size = 0
    with path.open('rb') as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
            size += len(block)
    st = path.stat()
    return {'path': str(path), 'bytes': size, 'mode': stat.S_IMODE(st.st_mode), 'sha256': h.hexdigest()}

def build_tree_digest(rows: list[dict]) -> str:
    h = hashlib.sha256()
    for row in rows:
        h.update(row['path'].encode())
        h.update(b'\0')
        h.update(str(row['bytes']).encode())
        h.update(b'\0')
        h.update(str(row['mode']).encode())
        h.update(b'\0')
        h.update(row['sha256'].encode())
        h.update(b'\n')
    return h.hexdigest()

def verify_source() -> dict:
    manifest = json.loads(MANIFEST.read_text())
    rows = sorted(manifest['files'], key=lambda row: row['path'])
    mismatches = []
    seen = set()
    actual_bytes = 0
    actual_rows = []
    for row in rows:
        rel = row['path']
        if rel in seen or Path(rel).is_absolute() or '..' in Path(rel).parts:
            raise RuntimeError('duplicate or unsafe manifest path: ' + rel)
        seen.add(rel)
        path = SOURCE / rel
        if not path.is_file() or path.is_symlink():
            mismatches.append({'path': rel, 'reason': 'missing_or_nonregular'})
            continue
        got = digest(path)
        actual_bytes += got['bytes']
        actual_rows.append({'path': rel, 'bytes': got['bytes'], 'mode': got['mode'], 'sha256': got['sha256']})
        for key in ('bytes', 'mode', 'sha256'):
            if got[key] != row[key]:
                mismatches.append({'path': rel, 'field': key, 'expected': row[key], 'actual': got[key]})
    actual_paths = set()
    for dp, dns, fns in os.walk(SOURCE, topdown=True, followlinks=False):
        dp = Path(dp)
        for name in dns:
            if (dp / name).is_symlink():
                raise RuntimeError('symlink source directory: ' + str(dp / name))
        for name in fns:
            path = dp / name
            if path.is_symlink() or not path.is_file():
                mismatches.append({'path': str(path), 'reason': 'nonregular_extra'})
            else:
                actual_paths.add(path.relative_to(SOURCE).as_posix())
    extra = sorted(actual_paths - seen)
    missing = sorted(seen - actual_paths)
    mismatches.extend({'path': x, 'reason': 'extra_file'} for x in extra)
    mismatches.extend({'path': x, 'reason': 'missing_file'} for x in missing)
    actual_rows.sort(key=lambda row: row['path'])
    manifest_desc = digest(MANIFEST)
    result = {
        'manifest': manifest_desc,
        'source_root': str(SOURCE),
        'file_count': len(rows),
        'actual_file_count': len(actual_rows),
        'manifest_bytes': manifest['bytes'],
        'actual_bytes': actual_bytes,
        'tree_digest_from_manifest': build_tree_digest(rows),
        'tree_digest_from_source': build_tree_digest(actual_rows),
        'mismatch_count': len(mismatches),
        'mismatches': mismatches,
        'verified': not mismatches and len(rows) == len(actual_rows) and actual_bytes == manifest['bytes'],
    }
    if not result['verified']:
        raise RuntimeError('source verification failed: ' + json.dumps(result, sort_keys=True))
    return result

def write_json(name: str, value: object) -> Path:
    path = RECEIPTS / name
    if path.exists():
        raise RuntimeError('refusing to overwrite receipt: ' + str(path))
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + '\n')
    return path

def binary_descriptor(path: Path) -> dict:
    result = {'path': str(path), 'exists': path.exists()}
    if not path.exists():
        return result
    result.update(digest(path))
    readelf = subprocess.run(['/usr/bin/readelf', '-n', str(path)], text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=False)
    build_id = None
    for line in readelf.stdout.splitlines():
        if 'Build ID:' in line:
            build_id = line.split('Build ID:', 1)[1].strip()
            break
    result.update({'readelf_returncode': readelf.returncode, 'build_id': build_id, 'readelf_sha256': hashlib.sha256(readelf.stdout.encode()).hexdigest()})
    return result

def main() -> int:
    if not SOURCE.is_dir() or not MANIFEST.is_file():
        raise SystemExit('missing frozen candidate source or manifest')
    if TARGET.exists():
        raise SystemExit('refusing existing target: ' + str(TARGET))
    if BUILD.exists():
        raise SystemExit('refusing existing build directory: ' + str(BUILD))
    for path in RETAINED_PATHS.values():
        if path.exists() or path.is_symlink():
            raise SystemExit('refusing existing retained executable: ' + str(path))
    for name in RECEIPT_NAMES:
        if (RECEIPTS / name).exists():
            raise SystemExit('refusing existing receipt: ' + str(RECEIPTS / name))
    for path in (TARGET, BUILD, RECEIPTS, LOGS, TMP, EXECUTABLES):
        path.mkdir(parents=True, exist_ok=True)
    source_start = verify_source()
    env = os.environ.copy()
    for key in REMOVED_FLAGS:
        env.pop(key, None)
    env.update(ENV_OVERRIDES)
    start = {
        'schema': 'litchi.performance.0964.candidate-managed-v3.build-start',
        'started_utc': utc(), 'source_root': str(SOURCE), 'source_manifest': source_start['manifest'],
        'source_verification_start': source_start, 'target_dir': str(TARGET), 'build_dir': str(BUILD),
        'cwd': str(SOURCE), 'command': COMMAND, 'env_overrides': ENV_OVERRIDES,
        'removed_environment_flags': REMOVED_FLAGS, 'deadline_seconds': DEADLINE_SECONDS,
        'roles': sorted(BINARY_PATHS), 'no_rustflags': True,
        'build_helper': digest(Path(__file__)),
        'rustc_identity': subprocess.check_output(['/home/zhuhe/.cargo/bin/rustc', '+1.95.0', '-vV'], text=True),
        'cargo_identity': subprocess.check_output(['/home/zhuhe/.cargo/bin/cargo', '+1.95.0', '-vV'], text=True),
        'effective_build_environment': {k: v for k, v in env.items() if k in ENV_OVERRIDES or k in REMOVED_FLAGS or k in ('CARGO_HOME', 'RUSTUP_HOME', 'RUSTUP_TOOLCHAIN', 'CARGO_BUILD_TARGET', 'CARGO_BUILD_RUSTC', 'CARGO_BUILD_RUSTDOC')},
    }
    start_path = write_json('candidate-managed-v3-build-start.json', start)
    stdout_log = BUILD / 'build.stdout.log'; stderr_log = BUILD / 'build.stderr.log'
    returncode = None; timed_out = False; launch_error = None; pid = None
    with stdout_log.open('w', encoding='utf-8') as stdout, stderr_log.open('w', encoding='utf-8') as stderr:
        try:
            proc = subprocess.Popen(COMMAND, cwd=SOURCE, env=env, stdout=stdout, stderr=stderr, start_new_session=True)
            pid = proc.pid
            try:
                returncode = proc.wait(timeout=DEADLINE_SECONDS)
            except subprocess.TimeoutExpired:
                timed_out = True
                try: os.killpg(proc.pid, signal.SIGTERM)
                except ProcessLookupError: pass
                try: returncode = proc.wait(timeout=20)
                except subprocess.TimeoutExpired:
                    try: os.killpg(proc.pid, signal.SIGKILL)
                    except ProcessLookupError: pass
                    returncode = proc.wait()
        except BaseException as exc:
            launch_error = repr(exc)
        stdout.flush(); stderr.flush()
    source_end = None; source_error = None
    try:
        source_end = verify_source()
    except BaseException as exc:
        source_error = repr(exc)
    terminal = {
        'schema': 'litchi.performance.0964.candidate-managed-v3.build-terminal',
        'finished_utc': utc(), 'pid': pid, 'process_group': pid, 'returncode': returncode,
        'timed_out': timed_out, 'launch_error': launch_error,
        'source_verification_end': source_end, 'source_verification_error': source_error,
        'stdout': digest(stdout_log), 'stderr': digest(stderr_log),
        'terminal_before_binary_parse': True,
    }
    terminal_path = write_json('candidate-managed-v3-build-terminal.json', terminal)
    binaries = {'schema':'litchi.performance.0964.candidate-managed-v3.build-binaries', 'terminal_receipt': digest(terminal_path), 'binaries': {}}
    for role, path in BINARY_PATHS.items():
        binaries['binaries'][role] = binary_descriptor(path)
    binary_path = write_json('candidate-managed-v3-build-binaries.json', binaries)
    retention = {'schema':'litchi.performance.0964.candidate-managed-v3.retention', 'terminal_receipt':digest(terminal_path), 'binary_receipt':digest(binary_path), 'binaries':{}}
    can_retain = returncode == 0 and not timed_out and launch_error is None and source_error is None and source_start == source_end and all(x.get('exists') for x in binaries['binaries'].values())
    if can_retain:
        for role, src in BINARY_PATHS.items():
            dst = RETAINED_PATHS[role]
            shutil.copy2(src, dst)
            srcd = binary_descriptor(src); dstd = binary_descriptor(dst)
            retention['binaries'][role] = {'build_output':srcd,'retained':dstd,'byte_match':all(srcd[k] == dstd[k] for k in ('bytes', 'mode', 'sha256'))}
        retention['status'] = all(x['byte_match'] for x in retention['binaries'].values())
    else:
        retention['status'] = False
        retention['reason'] = 'build/source verification did not satisfy retention preconditions'
    retention_path = write_json('candidate-managed-v3-retention.json', retention)
    result = {
        'schema':'litchi.performance.0964.candidate-managed-v3.build-result',
        'start_receipt':digest(start_path), 'terminal_receipt':digest(terminal_path),
        'binary_receipt':digest(binary_path), 'retention_receipt':digest(retention_path),
        'returncode':returncode, 'timed_out':timed_out, 'source_unchanged':source_start == source_end,
        'source_verification_start_pass':source_start['verified'], 'source_verification_end_pass':source_end is not None and source_end['verified'],
        'binary_parse_after_terminal':True, 'all_binaries_present':all(x.get('exists') for x in binaries['binaries'].values()),
        'retained_executables':retention['status'], 'launch_error':launch_error, 'source_error':source_error,
    }
    result_path = write_json('candidate-managed-v3-build-result.json', result)
    print(json.dumps({**result, 'result_receipt':digest(result_path)}, sort_keys=True))
    return 0 if result['returncode'] == 0 and result['source_unchanged'] and result['all_binaries_present'] and result['retained_executables'] else 1

if __name__ == '__main__':
    raise SystemExit(main())
