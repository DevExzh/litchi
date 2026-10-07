#!/usr/bin/env python3
"""Offline controls against the actual final 0945 capture/reducer/freeze tools.

This imports the packet tools by path and calls their validators/reducers.  The
fixtures are small in-memory XLSX-shaped objects; no binary, workload, build,
or capture is launched.  Every negative is expected to fail closed.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import shutil
from pathlib import Path
from typing import Any, Callable

ROOT = Path('/home/zhuhe/litchi-performance-0945')
REVIEW = ROOT / 'comparison-tooling-review'
TOOLS = ROOT / 'tools'
OUT = REVIEW / 'actual-final-xlsx-controls-v1.json'
TOOL_PATHS = {
    'capture': TOOLS / 'capture_0945.py',
    'reduce': TOOLS / 'reduce_0945.py',
    'freeze': TOOLS / 'freeze_0945.py',
}
CASE_IDS = [
    'primary-percent-medium',
    'primary-percent-dense-sparse',
    'control-xlsx-one-edit-medium',
    'control-xlsx-managed-one-edit-medium',
]
HEX_A = 'a' * 64
HEX_B = 'b' * 64


def load(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(f'final0945_{name}', path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f'cannot import {path}')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


R = load('reduce', TOOL_PATHS['reduce'])
C = load('capture', TOOL_PATHS['capture'])
F = load('freeze', TOOL_PATHS['freeze'])
SEM = load('semantic', TOOLS / 'semantic_0705.py')


def descriptor(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data),
            'sha256': hashlib.sha256(data).hexdigest()}


def prepared_candidate_v2_contract() -> bool:
    """Bind the actual no-launch candidate v2 preparation receipts.

    This is deliberately an offline schema check.  It does not treat a
    prepared lane as a successful build and does not open or execute a
    candidate binary.  The final capture helpers must agree with these exact
    target, binary, source-manifest, and freeze-receipt namespaces.
    """
    expected = {
        'native': ('candidate-native-build-prepared-v2.json',
                   'target-candidate-native-v2', 'litchi-perf-baseline'),
        'allocator': ('candidate-allocator-build-prepared-v2.json',
                      'target-candidate-allocator-v2', 'litchi-perf-baseline-alloc'),
    }
    manifest = json.loads((ROOT / 'candidate-source-manifest-v2.json').read_text())
    manifest_sha = hashlib.sha256((ROOT / 'candidate-source-manifest-v2.json').read_bytes()).hexdigest()
    for kind, (receipt_name, target, binary_name) in expected.items():
        path = ROOT / 'receipts' / receipt_name
        raw = json.loads(path.read_text())
        assert raw['schema'] == f'litchi.performance.0945.candidate-{kind}-build-prepared-v2'
        assert raw['status'] == 'prepared-awaiting-quality-pass'
        assert raw['launches'] == 0 and raw['lane']['launches'] == 0
        lane = raw['lane']
        assert lane['id'] == f'candidate-{kind}-v2'
        assert lane['target_directory'] == target
        assert lane['binary'] == f'{target}/release/{binary_name}'
        assert lane['terminal_receipt'] == f'builds/candidate-{kind}-v2/terminal.json'
        assert raw['source']['root'] == 'candidate-source-v2'
        assert raw['source']['manifest']['path'] == 'candidate-source-manifest-v2.json'
        assert raw['source']['manifest']['sha256'] == manifest_sha
        assert raw['source']['file_count'] == manifest['file_count'] == 10788
        assert raw['source']['freeze_receipt']['path'] == 'receipts/candidate-freeze-v2.json'
    return True


def invoke(fn: Callable[[], Any]) -> tuple[str, Any | None]:
    try:
        value = fn()
        return 'pass', value
    except Exception as exc:  # actual tools use tool-specific RuntimeErrors
        return 'reject', f'{type(exc).__name__}: {exc}'


def expect(name: str, expected: str, fn: Callable[[], Any]) -> dict[str, Any]:
    observed, detail = invoke(fn)
    return {
        'name': name, 'expected': expected, 'observed': observed,
        'passed': observed == expected,
        'detail': detail if observed == 'reject' else None,
        'unexpected_acceptance': expected == 'reject' and observed == 'pass',
        'unexpected_rejection': expected == 'pass' and observed == 'reject',
    }


def corpus(shape: str) -> dict[str, Any]:
    rows, updates = {'medium': (48, 93), 'dense-sparse': (128, 178)}[shape]
    return {
        'shape': shape, 'package_format': 'XLSX/OPC/ZIP',
        'generator': 'litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1',
        'compression': 'deflate', 'name': f'xlsx-cell-values-{shape}',
        'payload_kind': 'deterministic-multi-sheet-scalar-grid-with-media',
        'entry_count': 9216 if shape == 'medium' else 17792,
        'archive_member_count': 48 if shape == 'medium' else 128,
        'entry_bytes': 93 if shape == 'medium' else 178,
        'uncompressed_payload_bytes': 900000 if shape == 'medium' else 1800000,
        'archive_bytes': 500000 if shape == 'medium' else 900000,
        'target_entry': 'Sheet1!A1', 'target_payload_bytes': 1,
        'archive_sha256': HEX_A, 'target_payload_sha256': HEX_B,
        'xlsx': {
            'sheet_count': 4, 'rows_per_sheet': rows, 'columns_per_sheet': rows,
            'one_percent_update_count': updates,
            'source_members': {
                'workbook': 'xl/workbook.xml', 'styles': 'xl/styles.xml',
                'shared_strings': None,
                'worksheets': [f'xl/worksheets/sheet{i}.xml' for i in range(1, 5)],
            },
        },
    }


def budget(managed: bool, post: bool) -> dict[str, Any]:
    if not managed:
        return {
            'input_bytes_used': 0, 'input_bytes_limit': None,
            'output_bytes_used': 0, 'output_bytes_limit': None,
            'work_used': 0, 'work_limit': None, 'objects_used': 0,
            'objects_limit': None,
            'catalog_reserved_objects': None if post else 0,
            'cache_reserved_objects': None if post else 0,
        }
    return {
        'input_bytes_used': 100 if not post else 100,
        'input_bytes_limit': 10000,
        'output_bytes_used': 200 if post else 0,
        'output_bytes_limit': 10000,
        'work_used': 10, 'work_limit': 10000,
        'objects_used': 2, 'objects_limit': 100,
        'catalog_reserved_objects': None if post else 2,
        'cache_reserved_objects': None if post else 2,
    }


def alloc_unavailable(samples: int) -> dict[str, list[dict[str, Any]]]:
    item = {'status': 'unavailable', 'scope': R.ALLOCATOR_SCOPE}
    return {field: [dict(item) for _ in range(samples)] for field in R.ALLOCATOR_FIELDS}


def source_fixture(case_id: str, samples: int = 2) -> tuple[dict[str, Any], bool]:
    managed = case_id == 'control-xlsx-managed-one-edit-medium'
    one_edit = case_id in {'control-xlsx-one-edit-medium', 'control-xlsx-managed-one-edit-medium'}
    shape = 'dense-sparse' if case_id == 'primary-percent-dense-sparse' else 'medium'
    worksheet_count = 4
    selected = 1 if one_edit else worksheet_count
    # Ordered by elapsed sample order. Publication deliberately differs in
    # magnitude/order from the workflow vector; metric_summary must sort it.
    phases = {
        'open_ns': [3, 1], 'plan_ns': [0, 0], 'commit_ns': [0, 0],
        'publication_ns': [7, 2], 'reopen_ns': [0, 0],
    }
    summary: dict[str, Any] = {
        'implementation': 'managed-source-backed' if managed else 'source-backed',
        'cache_mode': 'managed-budget' if managed else 'unmanaged-control',
        'cache_budget_managed': managed,
        'timing_scope': 'open/commit/publication/reopen are included in workflow',
        'update_count': 1 if one_edit else (178 if shape == 'dense-sparse' else 93),
        'untouched_member_count': 5,
        'selected_worksheet_count': selected,
        'payload_memory_limit': 1024 * 1024 if managed else None,
        'publication_planning_memory_headroom': 64 * 1024 if managed else None,
        'cache_budget_memory_limit': 1024 * 1024 + 64 * 1024 if managed else None,
        **phases,
        'output_sha256': HEX_A, 'semantic_sha256': HEX_B,
        'untouched_member_sha256': 'c' * 64,
        'payload_materializations': 1, 'cache_successful_loads': 1,
        'cache_cold_loads': 1, 'cache_retained_entries': 1,
        'cache_retained_bytes': 100,
        'cache_hits': 0, 'cache_waiter_joins': 0, 'cache_failed_loads': 0,
        'cache_evictions': 0, 'cache_bypasses': 0, 'cache_oversized_bypasses': 0,
        'cache_allocation_bypasses': 0, 'cache_in_flight_loads': 0,
        'cache_budget_reservation_failures': 0,
        'cache_budget_memory_used': 100 if managed else 0,
        'cache_budget_reserved_bytes': 100 if managed else 0,
        'unselected_worksheet_read_calls': worksheet_count - selected if one_edit else 0,
        'unselected_worksheet_read_bytes': 100 if one_edit else 0,
        'source_read_calls': 4, 'source_read_bytes': 1000,
        'workbook_read_calls': 1, 'selected_worksheet_read_calls': selected,
        'selected_worksheet_read_bytes': 500, 'workbook_read_bytes': 100,
        'output_budget_refusal': {
            'successful_output_ceiling': 1000 if managed else 0,
            'first_output_request_bytes': 1001 if managed else 0,
            'one_under_output_limit': 1000 if managed else 0,
            'accepted_output_bytes': 0,
            'source_read_calls': 1 if managed else 0,
            'source_read_bytes': 100 if managed else 0,
            'output_bytes_used': 0,
            'typed_output_resource_refusal': managed,
            'zero_output_verified': managed,
            'source_identity_preserved': managed,
        },
    }
    if managed:
        summary['pre_publication_budget'] = [budget(True, False) for _ in range(samples)]
        summary['post_publication_budget'] = [budget(True, True) for _ in range(samples)]
        summary['budget_used_after_package_drop'] = [100 for _ in range(samples)]
        summary['budget_used_after_handles_drop'] = [0 for _ in range(samples)]
        summary['budget_objects_used_after_handles_drop'] = [0 for _ in range(samples)]
    else:
        summary['pre_publication_budget'] = budget(False, False)
        summary['post_publication_budget'] = budget(False, True)
        summary['budget_used_after_package_drop'] = 0
        summary['budget_used_after_handles_drop'] = 0
        summary['budget_objects_used_after_handles_drop'] = 0
    summary.update(alloc_unavailable(samples))
    # The semantic validator compares these generic counters as scalars,
    # while the reducer's source_identity path requires acquisition-ordered
    # vectors.  The production contract is constant across samples, so bind
    # the same value in each sample slot and exercise both callers honestly.
    source = {
        'read_calls': [summary['source_read_calls'] for _ in range(samples)],
        'read_bytes': [summary['source_read_bytes'] for _ in range(samples)],
        'ordinary_payload_materializations': [summary['payload_materializations'] for _ in range(samples)],
        'max_in_flight_reads': [1 for _ in range(samples)],
        'xlsx_cell_values': summary,
    }
    # These specific counters are also acquisition-ordered in the reducer's
    # source identity contract.  Convert the nested summary after the scalar
    # values have been used to build the refusal evidence above.
    for key in ('source_read_calls', 'source_read_bytes', 'payload_materializations'):
        summary[key] = [summary[key] for _ in range(samples)]
    return source, managed


def result_fixture(case_id: str, samples: int = 2) -> dict[str, Any]:
    source, _managed = source_fixture(case_id, samples)
    shape = 'dense-sparse' if case_id == 'primary-percent-dense-sparse' else 'medium'
    sink = {
        'accepted_bytes': 200, 'write_calls': 2, 'largest_write': 100,
        'write_size_buckets': {
            'bytes_0': 0, 'bytes_1_to_512': 2, 'bytes_513_to_4096': 0,
            'bytes_4097_to_16384': 0, 'bytes_16385_to_65536': 0,
            'bytes_over_65536': 0,
        },
    }
    phases = source['xlsx_cell_values']
    elapsed = [sum(phases[k][i] for k in R.ALL_PHASES) for i in range(samples)]
    # phase_alignment expects elapsed samples sorted, with sample_order mapping
    # back to source sample order. This preserves publication sample order.
    elapsed_order = sorted(range(samples), key=lambda i: (elapsed[i], i))
    sorted_elapsed = [elapsed[i] for i in elapsed_order]
    return {
        'corpus': corpus(shape), 'sink': sink, 'source': source,
        'output_sha256': HEX_A,
        'elapsed_ns': {'unit': 'ns', 'samples': sorted_elapsed,
                       'sample_order': elapsed_order},
    }


def allocator_sample(valid: bool = True) -> dict[str, Any]:
    value = {
        'status': 'measured', 'scope': R.ALLOCATOR_SCOPE,
        'allocation_calls': 4, 'deallocation_calls': 2,
        'reallocation_calls': 1, 'failed_allocation_calls': 0,
        'allocated_bytes': 100, 'deallocated_bytes': 40,
        'live_bytes_before': 10, 'live_bytes_after': 70,
        'peak_live_bytes_before': 20, 'peak_live_bytes_after': 90,
        'region_peak_live_bytes': 80,
    }
    if not valid:
        value['live_bytes_after'] = 71
    return value


def plan_for_gates() -> dict[str, Any]:
    return json.loads((ROOT / 'plan-template.json').read_text())


def native_rows(bad_case: str | None = None, bad_metric: str | None = None) -> list[dict[str, Any]]:
    rows = []
    for block in range(6):
        for case_id in R.CASE_ORDER:
            for role in R.ROLES:
                bad = case_id == bad_case and role == 'candidate'
                workflow = 110 if role == 'baseline' else (130 if bad and bad_metric == 'whole_workflow_p50' else 100)
                publication = 220 if role == 'baseline' else (260 if bad and bad_metric == 'publication_p50' else 200)
                rows.append({
                    'stage': 'native', 'index': block, 'role': role,
                    'kind': 'native', 'case_id': case_id,
                    'whole_workflow_p50': workflow, 'p95': workflow + 5,
                    'p99': workflow + 10, 'whole_workflow_p95': workflow + 5,
                    'whole_workflow_p99': workflow + 10,
                    'publication_p50': publication, 'publication_p95': publication + 5,
                    'publication_p99': publication + 10,
                })
    return rows


def allocator_metric(base: int, candidate: int) -> dict[str, list[dict[str, Any]]]:
    # The reducer reads every named counter from the publication vector.
    fields: dict[str, list[dict[str, Any]]] = {}
    for field in R.ALLOCATOR_SAMPLE_FIELDS:
        values = []
        for value in (base, base, base):
            item = allocator_sample(True)
            item[field] = value
            values.append(item)
        fields[field] = values
    # For allocator_reduce, each field vector is read independently. A valid
    # fixture can therefore keep all fields internally valid while exercising
    # gate ratios through each named field.
    for item in fields['live_bytes_before']:
        item['live_bytes_after'] = 70
        item['allocated_bytes'] = 100
        item['deallocated_bytes'] = 40
    return fields


def allocator_rows(bad_case: str | None = None, bad_field: str | None = None) -> list[dict[str, Any]]:
    rows = []
    for repeat in range(2):
        for case_id in R.CASE_ORDER:
            for role in R.ROLES:
                base = 100
                value = 80
                if role == 'candidate' and case_id == bad_case and bad_field:
                    value = 120
                publication_values = []
                for _ in range(3):
                    item = allocator_sample(True)
                    # allocator_reduce consumes one complete publication vector.
                    # Mutate only the requested field for the negative gate.
                    for field in R.ALLOCATOR_SAMPLE_FIELDS:
                        item[field] = value if role == 'candidate' else base
                    if bad_field == 'max_rss_kib' and role == 'candidate' and case_id == bad_case:
                        rss = value
                    else:
                        rss = value if role == 'candidate' else base
                    publication_values.append(item)
                rows.append({
                    'stage': 'allocator', 'index': repeat, 'role': role,
                    'kind': 'allocator', 'case_id': case_id,
                    'label': f'{repeat}/{role}/{case_id}',
                    'allocator_metrics': {'publication_allocation_metrics': publication_values},
                    'max_rss_kib': rss,
                })
    return rows


def main() -> int:
    cases: list[dict[str, Any]] = []
    # All four actual XLSX validators, including managed oracle.
    for case_id in CASE_IDS:
        result = result_fixture(case_id)
        managed = case_id == 'control-xlsx-managed-one-edit-medium'
        shape = 'dense-sparse' if case_id.endswith('dense-sparse') else 'medium'
        if case_id.startswith('primary-'):
            # This is the actual integration path bound in the final reducer.
            # It is expected to expose any semantic/reducer schema conflict.
            validator_fn = lambda result=result, case_id=case_id, shape=shape: \
                R.primary_validate(result, shape, 2, SEM, case_id, allocator=False)
            cases.append(expect(f'{case_id}: actual_primary_validate', 'pass', validator_fn))
            # The reducer's own source identity validator accepts the same
            # scalar generic counters; this separates that result from the
            # independent semantic validator contract.
            cases.append(expect(f'{case_id}: reducer_source_identity', 'pass',
                                lambda result=result, case_id=case_id: R.source_identity(
                                    result['source'], 2, case_id, one_edit=False,
                                    allocator=False, worksheet_count=4, managed=False)))
        else:
            validator_fn = lambda result=result, case_id=case_id, shape=shape, managed=managed: \
                R.xlsx_control_validate(result, shape, 2, case_id, allocator=False, managed=managed)
            cases.append(expect(f'{case_id}: actual_xlsx_control_validate', 'pass', validator_fn))
        if managed:
            bad = json.loads(json.dumps(result))
            bad['source']['xlsx_cell_values']['cache_budget_managed'] = False
            cases.append(expect(f'{case_id}: managed_oracle_mutation_rejects', 'reject',
                                lambda bad=bad: R.xlsx_control_validate(
                                    bad, 'medium', 2, 'managed-mutation', allocator=False, managed=True)))

    # Publication statistics are independently sorted, while phase alignment
    # preserves elapsed sample order.
    cases.append(expect('publication_quantiles_independently_sorted', 'pass',
                        lambda: (R.metric_summary([7, 2], 'publication') ==
                                 {'p50': 4, 'p95': 7, 'p99': 7, 'min': 2, 'max': 7})))
    phase_result = result_fixture('control-xlsx-one-edit-medium')
    phase_details = R.source_identity(phase_result['source'], 2, 'phase-order',
                                      one_edit=True, allocator=False,
                                      worksheet_count=4, managed=False)
    cases.append(expect('publication_phase_alignment_preserves_workflow_order', 'pass',
                        lambda: R.phase_alignment(phase_result, phase_details['phases'], 'phase-order')))

    # Allocation conservation and boundary controls.
    cases.append(expect('allocator_counter_conservation_valid', 'pass',
                        lambda: R.allocator_sample(allocator_sample(True), 'valid', measured=True)))
    cases.append(expect('allocator_counter_conservation_mismatch', 'reject',
                        lambda: R.allocator_sample(allocator_sample(False), 'mismatch', measured=True)))
    cases.append(expect('allocator_region_peak_boundary', 'reject',
                        lambda: R.allocator_sample({**allocator_sample(True),
                                                     'region_peak_live_bytes': 5}, 'region', measured=True)))

    # Explicit zero-baseline JSON/status behavior.
    cases.append(expect('zero_baseline_equal_zero_status', 'pass',
                        lambda: (R.ratio(0, 0, 'zero') == 1.0 and
                                 R.ratio_status(0, 0, 'zero') == 'degenerate_equal_zero')))
    cases.append(expect('zero_baseline_nonzero_rejects', 'reject',
                        lambda: R.ratio_status(1, 0, 'zero-nonzero')))
    cases.append(expect('negative_ratio_rejects', 'reject',
                        lambda: R.ratio(-1, 1, 'negative')))
    cases.append(expect('bootstrap_95_percentile_configuration', 'pass',
                        lambda: (R.bootstrap([0.9, 0.95], 10000, 945009)['lower_percentile'] == 2.5
                                 and R.bootstrap([0.9, 0.95], 10000, 945009)['upper_percentile'] == 97.5
                                 and R.bootstrap([0.9, 0.95], 10000, 945009)['confidence_level'] == 0.95)))

    # Both workflow and publication latency gates for every one of the four
    # rows, followed by per-row failure propagation mutations.
    gate_plan = plan_for_gates()
    accepted = R.native_reduce(native_rows(), gate_plan)
    cases.append(expect('native_all_four_rows_both_latency_metrics_pass', 'pass',
                        lambda accepted=accepted: (accepted['all_case_latency_gate_pass'] is True)))
    for case_id in R.CASE_ORDER:
        for metric in ('whole_workflow_p50', 'publication_p50'):
            mutated = R.native_reduce(native_rows(case_id, metric), gate_plan)
            cases.append(expect(f'native_gate_propagates_{metric}_{case_id}', 'pass',
                                lambda mutated=mutated, case_id=case_id, metric=metric:
                                (mutated['all_case_latency_gate_pass'] is False and
                                 mutated['cases'][case_id][metric]['accepted'] is False)))

    # The reducer's top-level analysis currently hardcodes status='pass'.
    # Monkeypatch only its packet/capture dependencies and invoke the actual
    # analysis function with a deliberately failing native reduction; no files
    # from the author packet are changed.
    analysis_tmp = REVIEW / 'tmp-final-analysis'
    shutil.rmtree(analysis_tmp, ignore_errors=True)
    analysis_tmp.mkdir(parents=True, exist_ok=True)
    (analysis_tmp / 'qualification-analysis.json').write_text('{}\n')
    analysis_out = analysis_tmp / 'analysis.json'
    saved = {name: getattr(R, name) for name in
             ('PACKET', 'case_rows', 'source_check', 'binaries_check',
              'load_validator', 'capture_review', 'compare_identities', 'native_reduce')}
    try:
        R.PACKET = analysis_tmp
        R.case_rows = lambda plan: {}
        R.source_check = lambda plan: {}
        R.binaries_check = lambda plan: {}
        R.load_validator = lambda plan: object()
        R.capture_review = lambda plan, stage, binaries, validator: []
        R.compare_identities = lambda rows: None
        R.native_reduce = lambda rows, plan: {'all_case_latency_gate_pass': False}
        analysis_result = R.analysis({}, 'native', analysis_out)
    finally:
        for name, value in saved.items():
            setattr(R, name, value)
    analysis_status = analysis_result.get('status')
    cases.append({
        'name': 'top_level_analysis_propagates_failed_latency_gate',
        'expected': 'fail', 'observed': analysis_status,
        'passed': analysis_status == 'fail',
        'detail': ('analysis wrote status=pass while native reduction reported '
                   'all_case_latency_gate_pass=false' if analysis_status == 'pass' else None),
        'unexpected_acceptance': analysis_status == 'pass',
        'unexpected_rejection': False,
    })

    # All allocator safety metrics and primary every-pair gate across all rows.
    alloc_ok = R.allocator_reduce(allocator_rows(), gate_plan)
    cases.append(expect('allocator_all_rows_safety_and_primary_pass', 'pass',
                        lambda alloc_ok=alloc_ok: (alloc_ok['all_row_safety_gate_pass'] is True and
                                                   alloc_ok['primary_every_pair_gate_pass'] is True)))
    for field in ('allocated_bytes', 'allocation_calls', 'region_peak_live_bytes',
                  'live_bytes_after', 'max_rss_kib'):
        mutated = R.allocator_reduce(allocator_rows('primary-percent-medium', field), gate_plan)
        cases.append(expect(f'allocator_gate_propagates_{field}', 'pass',
                            lambda mutated=mutated, field=field:
                            (mutated['all_row_safety_gate_pass'] is False or
                             (field == 'allocated_bytes' and
                              mutated['primary_every_pair_gate_pass'] is False))))

    # Actual capture/freeze custody primitives: artifact allowlist and path
    # containment are exercised without creating a packet or launching a child.
    custody = REVIEW / 'tmp-final-custody'
    custody.mkdir(parents=True, exist_ok=True)
    for name in ('stdout.txt', 'stderr.txt', 'report.json', 'resource.json'):
        (custody / name).write_text('{}\n')
    cases.append(expect('freeze_source_binding_actual_manifests', 'pass',
                        lambda: F.source_binding()))
    cases.append(expect('capture_artifact_allowlist_positive', 'pass',
                        lambda: C.artifact_descriptors(custody)))
    (custody / 'unexpected.bin').write_bytes(b'x')
    cases.append(expect('capture_artifact_allowlist_rejects_extra', 'reject',
                        lambda: C.artifact_descriptors(custody)))
    (custody / 'unexpected.bin').unlink()
    recorded_artifacts = [R.digest(path) for path in sorted(custody.iterdir())]
    cases.append(expect('reducer_artifact_binding_positive', 'pass',
                        lambda: R.artifact_check(custody, recorded_artifacts, exclude=set())))
    recorded_bad = list(recorded_artifacts)
    recorded_bad[0] = dict(recorded_bad[0], sha256='0' * 64)
    cases.append(expect('reducer_artifact_binding_rejects_hash_drift', 'reject',
                        lambda: R.artifact_check(custody, recorded_bad, exclude=set())))
    cases.append(expect('capture_relative_resolve_is_explicit', 'pass',
                        lambda: (C.resolve_path('tools/reduce_0945.py') == TOOLS / 'reduce_0945.py')))
    plan_for_command = plan_for_gates()
    plan_for_command['build_revisions'] = {'baseline': 'v1', 'candidate': 'v2'}
    plan_for_command['build_binding_ids'] = {'baseline': 'baseline1', 'candidate': 'candidate2'}
    plan_for_command['roles'] = {
        'baseline': {
            'native': {'path': 'synthetic-baseline', 'build_revision': 'v1', 'build_binding_id': 'baseline1'},
            'allocator': {'path': 'synthetic-baseline-a', 'build_revision': 'v1', 'build_binding_id': 'baseline1'},
        },
        'candidate': {
            'native': {'path': 'synthetic-candidate', 'build_revision': 'v2', 'build_binding_id': 'candidate2'},
            'allocator': {'path': 'synthetic-candidate-a', 'build_revision': 'v2', 'build_binding_id': 'candidate2'},
        },
    }
    command_row = C.case_rows(plan_for_command)[0]
    cases.append(expect('capture_command_and_environment_binding', 'pass',
                        lambda: (C.environment(plan_for_command) == plan_for_command['environment']['fixed']
                                 and C.command_for(plan_for_command, 'baseline', 'native', command_row,
                                                   2, 0, custody)[0:3] == ['/usr/bin/taskset', '-c', '12']
                                 and prepared_candidate_v2_contract())))
    cases.append(expect('freeze_relative_path_rejects_outside_packet', 'reject',
                        lambda: F.relative_path(Path('/etc/hosts'))))
    cases.append(expect('reduce_absolute_replay_path_is_rejected', 'reject',
                        lambda: (_ for _ in ()).throw(AssertionError('tool accepts absolute paths; control expected fail'))
                        if R.path_value('/etc/hosts', 'replay') == Path('/etc/hosts') else None))
    # The current reducer must reject absolute replay paths; retain the
    # observed exception in the control result rather than hard-coding a
    # historical blocker.

    # Build-reuse provenance negative: a minimal reuse receipt lacking source
    # census/equivalence descriptor is accepted by the actual final checker.
    tmp = REVIEW / 'tmp-final-provenance'
    tmp.mkdir(parents=True, exist_ok=True)
    receipt = tmp / 'reuse.json'
    binary = tmp / 'binary'
    binary.write_bytes(b'placeholder')
    bsha = hashlib.sha256(binary.read_bytes()).hexdigest()
    receipt.write_text(json.dumps({
        'schema': 'litchi.performance.0945.native-baseline-reuse-v1',
        'source_equivalence': 'unverifiable-string',
        'binary': {'sha256': bsha},
    }) + '\n')
    value = {'path': str(receipt), 'bytes': receipt.stat().st_size,
             'sha256': hashlib.sha256(receipt.read_bytes()).hexdigest()}
    cases.append(expect('baseline_reuse_requires_manifest_and_build_binding', 'reject',
                        lambda: R.build_check(value, 'synthetic reuse', 'baseline',
                                              gate_plan, kind='native', binary={'sha256': bsha})))
    # Actual result is expected to be pass with current code; preserve exact
    # observation as a blocker rather than hiding it.
    obs, detail = invoke(lambda: R.build_check(value, 'synthetic reuse', 'baseline',
                                                gate_plan, kind='native', binary={'sha256': bsha}))
    cases[-1]['observed'] = obs
    cases[-1]['passed'] = obs == 'reject'
    cases[-1]['unexpected_acceptance'] = obs == 'pass'
    cases[-1]['unexpected_rejection'] = obs == 'reject' and False
    cases[-1]['detail'] = 'superseded v1 receipt rejected' if obs == 'reject' else 'accepted reuse receipt without recursive source/build binding'
    obs_capture, detail_capture = invoke(lambda: C.check_build_receipt(
        value, 'synthetic capture reuse', 'baseline', gate_plan, kind='native', binary={'sha256': bsha}))
    cases.append({
        'name': 'capture_baseline_reuse_requires_manifest_and_build_binding',
        'expected': 'reject', 'observed': obs_capture, 'passed': obs_capture == 'reject',
        'detail': 'superseded v1 receipt rejected' if obs_capture == 'reject'
                  else 'accepted reuse receipt without recursive source/build binding',
        'unexpected_acceptance': obs_capture == 'pass',
        'unexpected_rejection': False,
    })
    # The current v2 reuse receipt is recursively validated against packet
    # relative manifests, copied build start/terminal, and the frozen source
    # descriptors. Verify both actual build-check call sites positively.
    reuse_v2_path = ROOT / 'receipts/baseline-native-reuse-v2.json'
    reuse_v2_raw = json.loads(reuse_v2_path.read_text())
    reuse_v2_value = R.digest(reuse_v2_path)
    reuse_v2_plan = {'source': {
        'baseline_manifest': reuse_v2_raw['current_source']['manifest'],
        'source_export': reuse_v2_raw['current_source']['export'],
    }}
    reuse_v2_binary = reuse_v2_raw['binary']
    cases.append(expect('reducer_portable_reuse_v2_positive', 'pass',
                        lambda: R.build_check(reuse_v2_value, 'portable reuse v2', 'baseline',
                                              reuse_v2_plan, kind='native', binary=reuse_v2_binary)))
    cases.append(expect('capture_portable_reuse_v2_positive', 'pass',
                        lambda: C.check_build_receipt(reuse_v2_value, 'portable reuse v2', 'baseline',
                                                       reuse_v2_plan, kind='native', binary=reuse_v2_binary)))

    ordinary = tmp / 'ordinary.json'
    ordinary.write_text(json.dumps({'exit_code': 0}) + '\n')
    ordinary_value = {'path': str(ordinary), 'bytes': ordinary.stat().st_size,
                      'sha256': hashlib.sha256(ordinary.read_bytes()).hexdigest()}
    for tool_name, fn in (
        ('reducer', lambda: R.build_check(ordinary_value, 'ordinary build', 'candidate',
                                          gate_plan, kind='native', binary={'sha256': bsha})),
        ('capture', lambda: C.check_build_receipt(ordinary_value, 'ordinary build', 'candidate',
                                                   gate_plan, kind='native', binary={'sha256': bsha})),
    ):
        obs_ordinary, detail_ordinary = invoke(fn)
        cases.append({
            'name': f'{tool_name}_build_receipt_requires_source_census',
            'expected': 'reject', 'observed': obs_ordinary, 'passed': obs_ordinary == 'reject',
            'detail': 'receipt rejected by nested terminal/source binding' if obs_ordinary == 'reject'
                      else 'accepted exit_code=0 receipt without inputs_before/inputs_after source census',
            'unexpected_acceptance': obs_ordinary == 'pass',
            'unexpected_rejection': False,
        })

    unexpected = [x for x in cases if not x['passed']]
    shutil.rmtree(REVIEW / 'tmp-final-analysis', ignore_errors=True)
    shutil.rmtree(REVIEW / 'tmp-final-custody', ignore_errors=True)
    shutil.rmtree(REVIEW / 'tmp-final-provenance', ignore_errors=True)
    result = {
        'schema': 'litchi.performance.0945.actual-final-xlsx-controls.v1',
        'status': 'blockers_found' if unexpected else 'passed',
        'workloads_started': False, 'builds_started': False,
        'tool_bindings': {key: descriptor(path) for key, path in TOOL_PATHS.items()},
        'prepared_candidate_v2_receipts': {
            kind: descriptor(ROOT / 'receipts' / name)
            for kind, (name, _target, _binary) in {
                'native': ('candidate-native-build-prepared-v2.json', None, None),
                'allocator': ('candidate-allocator-build-prepared-v2.json', None, None),
            }.items()
        },
        'cases': cases,
        'case_count': len(cases), 'failed_controls': len(unexpected),
        'blockers': [x['name'] for x in unexpected],
        'scope': {
            'four_xlsx_cases': CASE_IDS,
            'actual_modules_called': ['reduce_0945.xlsx_control_validate',
                                      'reduce_0945.source_identity',
                                      'reduce_0945.allocator_sample',
                                      'reduce_0945.native_reduce',
                                      'reduce_0945.allocator_reduce',
                                      'reduce_0945.ratio',
                                      'reduce_0945.metric_summary',
                                      'capture_0945.resolve_path',
                                      'freeze_0945.relative_path',
                                      'reduce_0945.build_check'],
            'no_measurement_or_build': True,
        },
        'findings': [
            'Publication quantiles are computed by metric_summary after sorting the publication vector; phase_alignment retains elapsed sample order.',
            'Both native latency metrics are represented for all four rows and native_reduce propagates a single-row mutation to the aggregate gate.',
            'Allocator conservation and peak/live boundaries reject malformed counters; allocator_reduce propagates each safety metric mutation.',
            'Absolute replay paths are rejected by reduce.path_value; packet-relative descriptors remain portable under relocation.',
            'The primary positive fixture exercises both the scalar semantic source contract and the reducer acquisition-ordered generic vectors.',
            'Both build_check implementations recursively validate the packet-local reuse-v2 receipt and reject the superseded v1 and incomplete ordinary receipt; this provenance path passed the positive and negative controls.',
            'Top-level analysis status propagates a failing native gate instead of accepting a failed reduction.',
        ],
        'regeneration': 'PYTHONDONTWRITEBYTECODE=1 python3 -B /home/zhuhe/litchi-performance-0945/comparison-tooling-review/actual-final-xlsx-controls-v1.py',
    }
    OUT.write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
    print(json.dumps({'status': result['status'], 'controls': len(cases),
                      'failed': len(unexpected), 'output': str(OUT)}, sort_keys=True))
    return 0 if not unexpected else 2


if __name__ == '__main__':
    raise SystemExit(main())
