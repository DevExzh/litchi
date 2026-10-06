#!/usr/bin/env python3
"""Offline publication ancestry addendum for the retained 0929 FP profile."""
from __future__ import annotations
import collections, hashlib, json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
import decode_fp_phase_ancestry_v1 as dec

ROOT = Path('/tmp/litchi-performance-0929')
LEAF = ROOT / 'capture-fp/perf-script-leaf-period.txt'
STACK = ROOT / 'capture-fp/perf-script-no-inline.txt'
MAP = ROOT / 'phase-map-fp-v1.json'
LOAD_BIAS = 0x5b8705390000
COMM = 'litchi-perf-bas'
DSO = '/tmp/litchi-performance-0929/target-fp/release/litchi-perf-baseline'


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda: f.read(1024 * 1024), b''):
            h.update(b)
    return h.hexdigest()


def join_rows():
    leaf, leaf_errors = dec.parse_export(LEAF, False)
    stack, stack_errors = dec.parse_export(STACK, True)
    sq = collections.defaultdict(collections.deque)
    for row in stack:
        sq[dec.event_key(row)].append(row)
    joined = []
    missing = 0
    for row in leaf:
        q = sq.get(dec.event_key(row))
        if not q:
            missing += 1
            continue
        other = q.popleft()
        row['frames'] = other['frames']
        row['stack_leaf_frame'] = other['leaf_frame']
        joined.append(row)
    return joined, leaf_errors, stack_errors, missing, sum(map(len, sq.values()))


def frame_symbols(row):
    return [frame['symbol'] for frame in row['frames'] if frame['symbol']]


def stats(rows):
    out = {'samples': len(rows), 'period': sum(r['period'] for r in rows)}
    return out


def add_stat(container, name, rows):
    container[name] = stats(rows)


def main():
    mapping = dec.load_mapping(MAP)
    joined, leaf_errors, stack_errors, missing_stack, missing_leaf = join_rows()
    publication = [
        r for r in joined
        if dec.classify(r, mapping, LOAD_BIAS, COMM, DSO)[0] == 'publication'
    ]
    inventory = [r for r in publication if any('Inventory::parse_with_limits' in s for s in frame_symbols(r))]
    ensure = [r for r in publication if any('ensure_source_section_inventory_xml' in s for s in frame_symbols(r))]
    union = [r for r in publication if (
        any('Inventory::parse_with_limits' in s for s in frame_symbols(r))
        or any('ensure_source_section_inventory_xml' in s for s in frame_symbols(r))
    )]
    both_inventory = [r for r in publication if (
        any('Inventory::parse_with_limits' in s for s in frame_symbols(r))
        and any('ensure_source_section_inventory_xml' in s for s in frame_symbols(r))
    )]
    inventory_union = [r for r in publication if (
        any('Inventory::parse_with_limits' in s for s in frame_symbols(r))
        or any('ensure_source_section_inventory_xml' in s for s in frame_symbols(r))
    )]
    publication_body_inventory = [r for r in inventory_union if any(
        'Package::publish_section_layout_commit_to_stream' in s for s in frame_symbols(r)
    )]
    publication_boundary_inventory = [r for r in inventory_union if r not in publication_body_inventory]
    sha = [r for r in publication if 'sha2::sha256::x86_sha::compress' in (r['leaf_frame']['symbol'] or '')]
    source_fp = [r for r in sha if any('litchi_opc::source_backed::SourceArtifact::fingerprint' in s for s in frame_symbols(r))]
    emit_hash = [r for r in sha if any('litchi_docx::source_backed::FingerprintingWriter' in s for s in frame_symbols(r))]
    both = [r for r in sha if r in source_fp and r in emit_hash]
    source_only = [r for r in source_fp if r not in emit_hash]
    emit_only = [r for r in emit_hash if r not in source_fp]
    neither = [r for r in sha if r not in source_fp and r not in emit_hash]
    out = {
        'schema': 'litchi.performance.0929.publication-attribution.v1',
        'status': 'descriptive_pass' if not (leaf_errors or stack_errors or missing_stack or missing_leaf) else 'failed_closed',
        'claim_scope': 'publication ancestry and leaf-period description only; inclusive rows are not additive and periods are sample-period weights, not elapsed time',
        'inputs': {
            'leaf_export': {'path': str(LEAF), 'bytes': LEAF.stat().st_size, 'sha256': sha256(LEAF)},
            'stack_export': {'path': str(STACK), 'bytes': STACK.stat().st_size, 'sha256': sha256(STACK)},
            'phase_map': {'path': str(MAP), 'bytes': MAP.stat().st_size, 'sha256': sha256(MAP)},
            'binary_sha256': 'c195a1d94acb7b240ca86ad32c13a9fd8b4c7b3affe085122970fe0f75c2169e',
            'source_sha256': '7904a6161e7b40a675d61acd9867b31b81a47f049b1ebb5a9979778bf14a373a',
            'load_bias': hex(LOAD_BIAS),
            'workload_comm': COMM,
            'workload_dso': DSO,
        },
        'join': {
            'leaf_rows': len(joined) + missing_stack,
            'stack_rows': len(joined) + missing_leaf,
            'joined_rows': len(joined),
            'missing_stack': missing_stack,
            'missing_leaf': missing_leaf,
            'leaf_parse_errors': leaf_errors,
            'stack_parse_errors': stack_errors,
        },
        'publication': stats(publication),
        'inventory_ancestry_inclusive': {},
        'inventory_operation_caller': {},
        'sha_leaf_caller_breakdown': {},
        'symbol_presence_notes': {
            'source_fingerprint_marker': 'litchi_opc::source_backed::SourceArtifact::fingerprint',
            'emitted_write_hash_marker': 'litchi_docx::source_backed::FingerprintingWriter<W> as std::io::Write',
            'sha_leaf_marker': 'sha2::sha256::x86_sha::compress',
            'caller_categories': 'Rows can contain both markers; source_only, emit_only, both, and neither are disjoint.',
        },
        'limits': [
            'The Inventory and ensure_source_section_inventory_xml counts are inclusive ancestry unions; nested frames are not summed.',
            'A SHA row is assigned by exact symbol presence in its retained FP callchain; absent caller symbols remain neither/unresolved.',
            'The profile includes correctness and publication setup work; these rows do not prove production latency or an optimization benefit.',
        ],
    }
    add_stat(out['inventory_ancestry_inclusive'], 'inventory_parse_with_limits', inventory)
    add_stat(out['inventory_ancestry_inclusive'], 'ensure_source_section_inventory_xml', ensure)
    add_stat(out['inventory_ancestry_inclusive'], 'inclusive_union', union)
    add_stat(out['inventory_ancestry_inclusive'], 'overlap_both', both_inventory)
    add_stat(out['inventory_operation_caller'], 'publication_body_with_publish_method', publication_body_inventory)
    add_stat(out['inventory_operation_caller'], 'mapped_publication_boundary_only', publication_boundary_inventory)
    add_stat(out['sha_leaf_caller_breakdown'], 'all_sha_leafs', sha)
    add_stat(out['sha_leaf_caller_breakdown'], 'source_artifact_fingerprint_inclusive', source_fp)
    add_stat(out['sha_leaf_caller_breakdown'], 'emitted_fingerprinting_writer_inclusive', emit_hash)
    add_stat(out['sha_leaf_caller_breakdown'], 'source_only', source_only)
    add_stat(out['sha_leaf_caller_breakdown'], 'emitted_only', emit_only)
    add_stat(out['sha_leaf_caller_breakdown'], 'both_markers', both)
    add_stat(out['sha_leaf_caller_breakdown'], 'neither_marker', neither)
    out['conservation'] = {
        'joined_period': sum(r['period'] for r in joined),
        'publication_period_share_of_joined': stats(publication)['period'] / sum(r['period'] for r in joined),
        'sha_caller_disjoint_period_sum': sum(stats(r)['period'] for r in (source_only, emit_only, both, neither)),
    }
    output = ROOT / 'publication-attribution-v1.json'
    output.write_text(json.dumps(out, indent=2, sort_keys=True) + '\n')
    print(json.dumps({'path': str(output), 'bytes': output.stat().st_size, 'sha256': sha256(output), 'status': out['status']}))

if __name__ == '__main__':
    main()
