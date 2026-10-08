#!/usr/bin/env python3
"""Validate the completed 0973 final-v2 replay and run one final-v2 noop replay."""
from __future__ import annotations
import datetime as dt
import hashlib
import importlib.util
import json
import os
import subprocess
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path('/home/zhuhe/litchi-performance-0976')
EXISTING = ROOT / 'final-v2-regression-0973-v1'
OUT = ROOT / 'final-v2-regression-0973-v2'
RECEIPT = ROOT / 'receipts/final-v2-regression-0973-v2-posthoc-v3.json'
RAW = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.perf.data')
BINARY = Path('/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile')
LAYOUT = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.target-layout.json')
COMMAND = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.command.json')
TERMINAL = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.terminal.json')
LEGACY_CORRECTED = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.corrected.perf-script-no-inline.txt')
STOCK = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.stock.perf-script-no-inline.txt')
NOOP = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.noop.perf-script-no-inline.txt')
LEGACY_NOOP_LOG = Path('/home/zhuhe/litchi-performance-0973/capture/instruction-profile-v2/child-01/child-01.noop.dwfl.jsonl')
PROTOCOL = Path('/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py')
SHIM = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.so'
SHIM_SOURCE = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.c'
SHIM_MAP = ROOT / 'tools/dwfl_dynamic_shim_0976_cache_final_v2.map'
FAILED_REPLAY_RECEIPT = ROOT / 'receipts/final-v2-regression-0973-v1-failure.json'
FAILED_NOOP_RECEIPT = ROOT / 'receipts/final-v2-regression-0973-v2-failure.json'
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
    spec = importlib.util.spec_from_file_location('protocol_common_0973_final_v2_noop', PROTOCOL)
    if spec is None or spec.loader is None:
        raise RuntimeError(PROTOCOL)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def parse_script(protocol: Any, path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
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
            'frames': [{'dso': frame['dso'], 'symbol_base': frame['symbol_base'],
                        'offset': frame['offset'], 'resolved': frame['resolved']}
                       for frame in sample['frames']],
        })
    public = {k: parsed[k] for k in (
        'sample_count', 'total_period', 'marker_sample_count', 'marker_period',
        'lost_events', 'empty_callchain_count', 'explicit_stack_truncated_count',
        'unresolved_frames',
    )}
    can_hash = hashlib.sha256(json.dumps(canonical, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    return public, headers, can_hash


def report_summary(path: Path, *, require_corrected: bool) -> dict[str, Any]:
    final = []
    with path.open(encoding='utf-8') as f:
        for line in f:
            row = json.loads(line)
            if row.get('schema') == 'litchi.performance.0976.dwfl_report.final':
                final.append(row)
    exact = [r for r in final if r.get('path_match') is True and r.get('name_match') is True]
    if require_corrected:
        fallback = [r for r in exact if r.get('admission') is not True]
        stored = [r for r in exact if r.get('cache_stored') is True]
        reused = [r for r in exact if r.get('reused_existing') is True]
        checks = {
            'exact_rows_nonzero': bool(exact), 'all_exact_admitted': all(r.get('admission') is True for r in exact),
            'fallback_exact_rows_zero': not fallback, 'one_cache_store': len(stored) == 1,
            'remaining_rows_reused': len(reused) == len(exact) - 1,
            'all_identity_valid': all(r.get('identity_valid') is True for r in exact),
            'all_config_valid': all(r.get('config_valid') is True for r in exact),
            'all_result_modules_non_nil': all(str(r.get('result_module')) != '(nil)' for r in exact),
        }
        if not all(checks.values()):
            raise AssertionError(f'corrected final-v2 report checks failed: {checks}')
        result = {'exact_target_rows': len(exact), 'cache_stored_rows': len(stored),
                  'reused_existing_rows': len(reused), 'fallback_exact_rows': len(fallback),
                  'identity_valid_rows': len([r for r in exact if r.get('identity_valid') is True]),
                  'config_valid_rows': len([r for r in exact if r.get('config_valid') is True]),
                  'all_final_report_rows': len(final), 'non_target_final_report_rows': len(final)-len(exact),
                  'checks': checks}
    else:
        unchanged = [r for r in exact if r.get('forwarded_base') == r.get('original_base')]
        changed = [r for r in exact if r.get('forwarded_base') != r.get('original_base')]
        checks = {
            'exact_rows_nonzero': bool(exact), 'all_exact_rejected': all(r.get('admission') is False for r in exact),
            'all_exact_forwarded_unchanged': len(unchanged) == len(exact),
            'no_cache_store': not any(r.get('cache_stored') is True for r in exact),
            'no_reuse': not any(r.get('reused_existing') is True for r in exact),
            'no_base_changes': not changed,
        }
        if not all(checks.values()):
            raise AssertionError(f'noop final-v2 report checks failed: {checks}')
        result = {'exact_target_rows': len(exact), 'forwarded_unchanged_rows': len(unchanged),
                  'base_changed_rows': len(changed), 'all_final_report_rows': len(final),
                  'non_target_final_report_rows': len(final)-len(exact), 'checks': checks}
    result['file'] = binding(path)
    result['mode_counts'] = dict(Counter(r.get('mode') for r in final))
    return result


def main() -> int:
    if RECEIPT.exists():
        raise FileExistsError('refusing overwrite')
    required = [EXISTING/'child-01.final-v2.perf-script-no-inline.txt', EXISTING/'child-01.final-v2.dwfl.jsonl',
                RAW, BINARY, LAYOUT, COMMAND, TERMINAL, LEGACY_CORRECTED, STOCK, NOOP, LEGACY_NOOP_LOG,
                PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP, FAILED_REPLAY_RECEIPT, FAILED_NOOP_RECEIPT]
    for p in required:
        if not p.is_file(): raise FileNotFoundError(p)
    inputs = [RAW, BINARY, LAYOUT, COMMAND, TERMINAL, LEGACY_CORRECTED, STOCK, NOOP, LEGACY_NOOP_LOG,
              PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP, FAILED_REPLAY_RECEIPT, FAILED_NOOP_RECEIPT]
    before = {str(p): binding(p) for p in inputs}
    protocol = load_protocol()
    legacy_noop_first = json.loads(LEGACY_NOOP_LOG.open(encoding='utf-8').readline())
    if legacy_noop_first.get('mode') != 'noop':
        raise AssertionError('retained noop log is not marked mode=noop')
    final_summary, final_headers, final_hash = parse_script(protocol, EXISTING/'child-01.final-v2.perf-script-no-inline.txt')
    legacy_summary, legacy_headers, legacy_hash = parse_script(protocol, LEGACY_CORRECTED)
    stock_summary, stock_headers, stock_hash = parse_script(protocol, STOCK)
    noop_summary, noop_headers, noop_hash = parse_script(protocol, NOOP)
    recovery = {
        'sample_count_6931': final_summary['sample_count'] == 6931,
        'period_6931000000': final_summary['total_period'] == 6_931_000_000,
        'known_historical_lost_events_8': final_summary['lost_events'] == 8,
        'legacy_lost_events_match': final_summary['lost_events'] == legacy_summary['lost_events'] == 8,
        'empty_callchains_zero': final_summary['empty_callchain_count'] == 0,
        'marker_count_947': final_summary['marker_sample_count'] == 947,
        'marker_period_947000000': final_summary['marker_period'] == 947_000_000,
    }
    if not all(recovery.values()): raise AssertionError(f'0973 final recovery checks failed: {recovery}')
    controls = {
        'final_vs_legacy_headers': final_headers == legacy_headers,
        'final_vs_stock_headers': final_headers == stock_headers,
        'final_vs_noop_headers': final_headers == noop_headers,
        'final_vs_legacy_canonical': final_hash == legacy_hash,
        'stock_noop_bytes': sha256(STOCK) == sha256(NOOP),
        'stock_noop_headers': stock_headers == noop_headers,
    }
    if not all(controls.values()): raise AssertionError(f'0973 identity controls failed: {controls}')
    corrected_reports = report_summary(EXISTING/'child-01.final-v2.dwfl.jsonl', require_corrected=True)
    # The final-v2 noop replay already completed in v1 and is immutable input
    # here. Record its exact command contract, but do not invoke perf again.
    output_script = OUT/'child-01.final-v2-noop.perf-script-no-inline.txt'
    output_stderr = OUT/'child-01.final-v2-noop.perf.stderr'
    output_log = OUT/'child-01.final-v2-noop.dwfl.jsonl'
    layout = json.loads(LAYOUT.read_text())
    env = os.environ.copy()
    env.update({'LC_ALL':'C','LANG':'C','TZ':'UTC','LITCHI_0973_SHIM_MODE':'noop',
                'LITCHI_0973_SHIM_LOG':str(output_log),'LITCHI_0973_TARGET_PATH':TARGET_PATH,
                'LITCHI_0973_TARGET_NAME':TARGET_NAME,'LITCHI_0973_CORRECTED_BASE':layout['corrected_base'],
                'LITCHI_0973_PERF_EXPECTED_START':layout['perf_expected_start'],
                'LITCHI_0976_TARGET_BUILD_ID':TARGET_BUILD_ID,'LITCHI_0976_TARGET_SHA256':TARGET_SHA256,
                'LITCHI_0976_EXPECTED_DELTA':'0x1000','LD_PRELOAD':str(SHIM)})
    argv = ['/usr/bin/perf','script','--show-lost-events','--no-inline','-F','cpu,pid,tid,time,event,period,ip,sym,symoff,dso','-i',str(RAW)]
    if not output_script.is_file() or not output_stderr.is_file() or not output_log.is_file():
        raise FileNotFoundError('completed final-v2 noop output')
    noop_final_summary, noop_final_headers, noop_final_hash = parse_script(protocol, output_script)
    noop_recovery = {
        'sample_count_6931': noop_final_summary['sample_count'] == 6931,
        'period_6931000000': noop_final_summary['total_period'] == 6_931_000_000,
        'known_historical_lost_events_8': noop_final_summary['lost_events'] == 8,
        'empty_callchains_match_stock': noop_final_summary['empty_callchain_count'] == stock_summary['empty_callchain_count'],
        'marker_count_match_stock': noop_final_summary['marker_sample_count'] == stock_summary['marker_sample_count'],
        'marker_period_match_stock': noop_final_summary['marker_period'] == stock_summary['marker_period'],
        'unresolved_match_stock': noop_final_summary['unresolved_frames'] == stock_summary['unresolved_frames'],
    }
    if not all(noop_recovery.values()): raise AssertionError(f'0973 noop recovery failed: {noop_recovery}')
    noop_reports = report_summary(output_log, require_corrected=False)
    noop_controls = {'noop_output_equals_stock': sha256(output_script) == sha256(STOCK),
                     'noop_output_equals_retained_noop': sha256(output_script) == sha256(NOOP),
                     'noop_headers_equal_stock': noop_final_headers == stock_headers,
                     'noop_canonical_equals_stock': noop_final_hash == stock_hash,
                     'noop_stderr_empty': output_stderr.stat().st_size == 0}
    if not all(noop_controls.values()): raise AssertionError(f'final-v2 noop identity failed: {noop_controls}')
    after = {str(p): binding(p) for p in inputs}
    if after != before: raise AssertionError('bound inputs changed during validation/noop replay')
    receipt = {'schema':'litchi.performance.0976.final_v2_regression_0973.posthoc.v2','status':'pass','created_utc':utc(),
      'purpose':'posthoc validation of retained 0973 final-v2 corrected and completed final-v2 noop replays; no decoder rerun, workload, or build',
      'inputs':before,'prior_failed_replay_receipt':binding(FAILED_REPLAY_RECEIPT),'prior_failed_noop_receipt':binding(FAILED_NOOP_RECEIPT),
      'final_v2_binding':{'shim':binding(SHIM),'source':binding(SHIM_SOURCE),'version_script':binding(SHIM_MAP),'target_build_id':TARGET_BUILD_ID,'target_sha256':TARGET_SHA256},
      'corrected_replay':{'output':binding(EXISTING/'child-01.final-v2.perf-script-no-inline.txt'),'log':binding(EXISTING/'child-01.final-v2.dwfl.jsonl'),'summary':final_summary,'canonical_sha256':final_hash,'recovery':recovery,'report':corrected_reports},
      'retained_legacy_noop_log_first_row': legacy_noop_first,
      'retained_controls':{'legacy_corrected':{'binding':binding(LEGACY_CORRECTED),'summary':legacy_summary,'canonical_sha256':legacy_hash},'stock':{'binding':binding(STOCK),'summary':stock_summary,'canonical_sha256':stock_hash},'noop':{'binding':binding(NOOP),'summary':noop_summary,'canonical_sha256':noop_hash},'header_and_canonical_controls':controls},
      'noop_replay':{'argv':argv,'environment':{k:env[k] for k in ('LITCHI_0973_SHIM_MODE','LITCHI_0973_TARGET_PATH','LITCHI_0973_TARGET_NAME','LITCHI_0973_CORRECTED_BASE','LITCHI_0973_PERF_EXPECTED_START','LITCHI_0976_TARGET_BUILD_ID','LITCHI_0976_TARGET_SHA256','LITCHI_0976_EXPECTED_DELTA','LD_PRELOAD')},'returncode':0,'reused_existing_output':True,'script':binding(output_script),'stderr':binding(output_stderr),'log':binding(output_log),'summary':noop_final_summary,'canonical_sha256':noop_final_hash,'recovery':noop_recovery,'report':noop_reports,'identity_controls':noop_controls},
      'scope':{'stock_scope':'stock was retained no-preload output; it was not run with final_v2.so.','legacy_noop_scope':'retained noop output carries the historical 0973 shim mode=noop (verified from its first log row; old shim, not final_v2.so); its binding is retained for comparison.','final_noop_scope':'the preserved noop replay is the exact final_v2.so noop control; this posthoc validator does not rerun it.','loss':'eight lost events are a known property of the historical 0973 raw capture; no zero-loss claim is made.'},
      'native_workload_started':False,'native_capture_started':False,'native_build_started':False,'target_executed':False}
    RECEIPT.write_text(json.dumps(receipt,indent=2,sort_keys=True)+'\n')
    print(json.dumps({'status':'pass','receipt':str(RECEIPT),'receipt_sha256':sha256(RECEIPT),'recovery':recovery,'noop_recovery':noop_recovery,'controls':controls,'noop_controls':noop_controls,'final_hash':final_hash,'noop_hash':noop_final_hash,'corrected_reports':corrected_reports,'noop_reports':noop_reports},sort_keys=True))
    return 0

if __name__ == '__main__': raise SystemExit(main())
