#!/usr/bin/env python3
"""One offline final-v2 replay and header controls for retained 0973 child-01."""
from __future__ import annotations
import datetime as dt
import hashlib
import importlib.util
import json
import os
import subprocess
from pathlib import Path
from typing import Any

ROOT = Path('/home/zhuhe/litchi-performance-0976')
OUT = ROOT / 'final-v2-regression-0973-v1'
RECEIPT = ROOT / 'receipts/final-v2-regression-0973-v1.json'
RAW = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.perf.data')
BINARY = Path('/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile')
LAYOUT = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.target-layout.json')
COMMAND = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.command.json')
TERMINAL = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.terminal.json')
LEGACY_CORRECTED = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.corrected.perf-script-no-inline.txt')
STOCK = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.stock.perf-script-no-inline.txt')
NOOP = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.noop.perf-script-no-inline.txt')
PROTOCOL = Path('/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py')
SHIM = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.so'
SHIM_SOURCE = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.c'
SHIM_MAP = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.map'
TARGET_PATH = str(BINARY.resolve())
TARGET_NAME = BINARY.name
TARGET_BUILD_ID = 'e5c457384eeebd057932c444efda77618f137399'
TARGET_SHA256 = '0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb'
EXPECTED_PID = 3726714
EXPECTED_CPU = 31
PERIOD = 1_000_000


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {'path': str(path.resolve()), 'bytes': path.stat().st_size, 'sha256': sha256(path)}


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load_protocol() -> Any:
    spec = importlib.util.spec_from_file_location('protocol_common_0973_final_v2_regression', PROTOCOL)
    if spec is None or spec.loader is None:
        raise RuntimeError(PROTOCOL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def parsed_summary(protocol: Any, path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    parsed = protocol.parse_perf_script_strict(
        path.read_text(encoding='utf-8'), marker_dso=TARGET_PATH,
        expected_pid=EXPECTED_PID, expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
    )
    headers = []
    canonical = []
    for sample in parsed['samples']:
        headers.append({k: sample[k] for k in ('pid', 'cpu', 'time', 'period', 'event')})
        canonical.append({
            'pid': sample['pid'], 'cpu': sample['cpu'], 'time': sample['time'],
            'period': sample['period'], 'event': sample['event'],
            'frames': [{
                'dso': frame['dso'], 'symbol_base': frame['symbol_base'],
                'offset': frame['offset'], 'resolved': frame['resolved'],
            } for frame in sample['frames']],
        })
    public = {k: parsed[k] for k in (
        'sample_count', 'total_period', 'marker_sample_count', 'marker_period',
        'lost_events', 'empty_callchain_count', 'explicit_stack_truncated_count',
        'unresolved_frames',
    )}
    canonical_hash = hashlib.sha256(json.dumps(canonical, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    return public, headers, canonical_hash


def log_summary(path: Path) -> dict[str, Any]:
    final = []
    with path.open(encoding='utf-8') as f:
        for line in f:
            row = json.loads(line)
            if row.get('schema') == 'litchi.performance.0976.dwfl_report.final':
                final.append(row)
    exact = [r for r in final if r.get('path_match') is True and r.get('name_match') is True]
    fallback = [r for r in exact if r.get('admission') is not True]
    stored = [r for r in exact if r.get('cache_stored') is True]
    reused = [r for r in exact if r.get('reused_existing') is True]
    checks = {
        'exact_rows_nonzero': bool(exact),
        'all_exact_admitted': all(r.get('admission') is True for r in exact),
        'fallback_exact_rows_zero': not fallback,
        'one_cache_store': len(stored) == 1,
        'remaining_rows_reused': len(reused) == len(exact) - 1,
        'all_identity_valid': all(r.get('identity_valid') is True for r in exact),
        'all_config_valid': all(r.get('config_valid') is True for r in exact),
        'all_result_modules_non_nil': all(str(r.get('result_module')) != '(nil)' for r in exact),
    }
    if not all(checks.values()):
        raise AssertionError(f'final-v2 log checks failed: {checks}')
    return {
        'file': binding(path), 'all_final_report_rows': len(final),
        'non_target_final_report_rows': len(final) - len(exact),
        'exact_target_rows': len(exact), 'cache_stored_rows': len(stored),
        'reused_existing_rows': len(reused), 'fallback_exact_rows': len(fallback),
        'identity_valid_rows': sum(r.get('identity_valid') is True for r in exact),
        'config_valid_rows': sum(r.get('config_valid') is True for r in exact),
        'errno_after_counts': {str(k): v for k, v in __import__('collections').Counter(r.get('dwfl_errno_after') for r in exact).items()},
        'checks': checks,
    }


def main() -> int:
    if OUT.exists() or RECEIPT.exists():
        raise FileExistsError('refusing overwrite of regression output or receipt')
    inputs = [RAW, BINARY, LAYOUT, COMMAND, TERMINAL, LEGACY_CORRECTED, STOCK, NOOP,
              PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP]
    for path in inputs:
        if not path.is_file():
            raise FileNotFoundError(path)
    before = {str(path): binding(path) for path in inputs}
    layout = json.loads(LAYOUT.read_text(encoding='utf-8'))
    if layout['binary']['sha256'] != TARGET_SHA256 or layout['executable_mapping']['build_id'] != TARGET_BUILD_ID:
        raise AssertionError('0973 layout binary/build ID custody mismatch')
    if layout['corrected_base'] != '0x563291817000' or layout['perf_expected_start'] != '0x563291818000':
        raise AssertionError('0973 layout base contract changed')
    OUT.mkdir(parents=True)
    script = OUT / 'child-01.final-v2.perf-script-no-inline.txt'
    stderr = OUT / 'child-01.final-v2.perf.stderr'
    log = OUT / 'child-01.final-v2.dwfl.jsonl'
    env = os.environ.copy()
    env.update({
        'LC_ALL': 'C', 'LANG': 'C', 'TZ': 'UTC',
        'LITCHI_0973_SHIM_MODE': 'corrected', 'LITCHI_0973_SHIM_LOG': str(log),
        'LITCHI_0973_TARGET_PATH': TARGET_PATH, 'LITCHI_0973_TARGET_NAME': TARGET_NAME,
        'LITCHI_0973_CORRECTED_BASE': layout['corrected_base'],
        'LITCHI_0973_PERF_EXPECTED_START': layout['perf_expected_start'],
        'LITCHI_0976_TARGET_BUILD_ID': TARGET_BUILD_ID,
        'LITCHI_0976_TARGET_SHA256': TARGET_SHA256,
        'LITCHI_0976_EXPECTED_DELTA': '0x1000', 'LD_PRELOAD': str(SHIM),
    })
    argv = ['/usr/bin/perf', 'script', '--show-lost-events', '--no-inline',
            '-F', 'cpu,pid,tid,time,event,period,ip,sym,symoff,dso', '-i', str(RAW)]
    with script.open('wb') as out, stderr.open('wb') as err:
        run = subprocess.run(argv, stdout=out, stderr=err, env=env, stdin=subprocess.DEVNULL, check=False, timeout=600)
    if run.returncode != 0 or not log.is_file():
        raise RuntimeError(f'final-v2 0973 replay failed rc={run.returncode}')
    protocol = load_protocol()
    final_summary, final_headers, final_hash = parsed_summary(protocol, script)
    legacy_summary, legacy_headers, legacy_hash = parsed_summary(protocol, LEGACY_CORRECTED)
    stock_summary, stock_headers, stock_hash = parsed_summary(protocol, STOCK)
    noop_summary, noop_headers, noop_hash = parsed_summary(protocol, NOOP)
    recovery = {
        'sample_count_matches_retained_6931': final_summary['sample_count'] == 6931,
        'period_matches_retained_6931000000': final_summary['total_period'] == 6_931_000_000,
        'lost_events_zero': final_summary['lost_events'] == 0,
        'empty_callchains_zero': final_summary['empty_callchain_count'] == 0,
    }
    if not all(recovery.values()):
        raise AssertionError(f'0973 final-v2 recovery checks failed: {recovery}')
    headers_equal = {
        'final_vs_legacy_corrected': final_headers == legacy_headers,
        'final_vs_stock': final_headers == stock_headers,
        'final_vs_noop': final_headers == noop_headers,
        'stock_vs_noop_bytes': sha256(STOCK) == sha256(NOOP),
    }
    if not all(headers_equal.values()):
        raise AssertionError(f'0973 sample/header identity controls failed: {headers_equal}')
    report = log_summary(log)
    after = {str(path): binding(path) for path in inputs}
    if after != before:
        raise AssertionError('bound 0973 inputs changed during final-v2 replay')
    receipt = {
        'schema': 'litchi.performance.0976.final_v2_regression_0973.v1',
        'status': 'pass', 'created_utc': utc(),
        'purpose': 'one offline final-v2 decoder replay of retained 0973 child-01 with immutable stock/noop/header controls; no workload or build',
        'inputs': before,
        'final_v2_binding': {'shim': binding(SHIM), 'source': binding(SHIM_SOURCE), 'version_script': binding(SHIM_MAP),
                             'target_build_id': TARGET_BUILD_ID, 'target_sha256': TARGET_SHA256,
                             'environment_contract': {k: env[k] for k in ('LITCHI_0973_SHIM_MODE','LITCHI_0973_TARGET_PATH','LITCHI_0973_TARGET_NAME','LITCHI_0973_CORRECTED_BASE','LITCHI_0973_PERF_EXPECTED_START','LITCHI_0976_TARGET_BUILD_ID','LITCHI_0976_TARGET_SHA256','LITCHI_0976_EXPECTED_DELTA','LD_PRELOAD')}},
        'command': {'argv': argv, 'returncode': run.returncode, 'stdout': binding(script), 'stderr': binding(stderr), 'dwfl_log': binding(log)},
        'final_v2_summary': final_summary,
        'final_v2_canonical_sha256': final_hash,
        'retained_summaries': {'legacy_corrected': legacy_summary, 'stock': stock_summary, 'noop': noop_summary},
        'retained_canonical_sha256': {'legacy_corrected': legacy_hash, 'stock': stock_hash, 'noop': noop_hash},
        'recovery_checks': recovery,
        'header_identity_controls': headers_equal,
        'stock_noop_scope': 'stock/noop files are immutable no-preload controls; they are not claimed to have run with final_v2.so. Their sample headers/time/periods are compared with the final_v2 output.',
        'scope_limits': ['0973 has 6931 samples and is a historical control, not the 0975 6940-sample recovery gate.', 'No performance or speedup claim follows.'],
        'native_workload_started': False, 'native_capture_started': False, 'native_build_started': False, 'target_executed': False,
    }
    RECEIPT.write_text(json.dumps(receipt, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print(json.dumps({'status':'pass','receipt':str(RECEIPT),'receipt_sha256':sha256(RECEIPT),'final_summary':final_summary,'final_canonical_sha256':final_hash,'report':report,'header_identity_controls':headers_equal},sort_keys=True))
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
