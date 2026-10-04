"""Fail-closed reduction of the generated-fixture owning-input qualification."""
import argparse
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
SOURCE_SHA = '61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757'
OUTPUT_SHA = '23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9'
POSITIONS = [0, 28, 57, 85, 114, 142, 171, 199]
SELECTED = sorted(f'ppt/slides/slide{i + 1}.xml' for i in POSITIONS)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def desc(path):
    data = path.read_bytes()
    return {'path': str(path), 'bytes': len(data), 'sha256': sha(data)}


def check(row):
    path = Path(row['path'])
    assert desc(path) == {k: row[k] for k in ('path', 'bytes', 'sha256')}, path


def load(name):
    return json.loads((P / name).read_text())


def reduce():
    capture = load('capture-v3/command.json')
    build = load('build-v3/command.json')
    assert capture['exit_code'] == build['exit_code'] == 0
    for receipt in (capture, build):
        check(receipt['helper'])
        check(receipt['source'])
        for row in receipt['logs']:
            check(row)
    for row in capture['artifacts']:
        check(row)
    binary = load('binary.json')['retained']
    check(binary)
    assert binary['sha256'] == build['binary']['sha256'] == capture['binary']['sha256']
    quality = load('quality-v3-v2/receipt.json')
    assert quality['status'] == 'pass' and len(quality['commands']) == 2
    assert all(item['exit_code'] == 0 for item in quality['commands'])
    report = load('capture-v3/artifacts/report.json')
    assert report['status'] == 'probe_passed_pending_independent_raw_census'
    gates = report['gates']
    assert gates == load('capture-v3/artifacts/gates.json')
    for key in ('source_identity_verified', 'source_shape_verified', 'owning_ingress_verified',
                'exact_noop_verified', 'semantic_output_verified', 'untouched_logical_parts_verified',
                'forward_output_verified', 'inverse_exact_source_verified',
                'forward_replay_exact_output_verified', 'durable_forward_roundtrip_verified',
                'durable_inverse_reopen_verified', 'finite_input_limit_verified',
                'source_immutability_verified', 'raw_census_artifacts_exported'):
        assert gates.get(key) is True, key
    assert report['selected_slide_positions'] == POSITIONS and report['replacement_count'] == 64
    assert report['slides'] == 200 and report['text_boxes_per_slide'] == report['media_entries'] == 8
    assert report['corpus']['entry_count'] == 229 and report['corpus']['archive_member_count'] == 445
    assert report['ownership'] == {
        'ingress_unmodified_owned_source': True, 'ingress_exact_source_len': 17017139,
        'ingress_exact_source_sha256': SOURCE_SHA, 'after_edit_unmodified_owned_source': False,
        'after_edit_retained_source_limits': True, 'after_edit_exact_source_len': None,
        'after_edit_exact_source_sha256': None, 'after_noop_commit_unmodified_owned_source': True,
    }
    for key, expected in [('source', SOURCE_SHA), ('no_op', SOURCE_SHA), ('inverse', SOURCE_SHA), ('output', OUTPUT_SHA), ('forward_replay', OUTPUT_SHA)]:
        row = report[key]
        data = (P / 'capture-v3/artifacts' / row['file']).read_bytes()
        assert len(data) == row['bytes'] and sha(data) == row['sha256'] == expected
    for key in ('stale_source_refusal', 'foreign_source_refusal'):
        assert gates[key]['refused'] and gates[key]['output_unchanged']
        assert gates[key]['error_debug'] == 'UnsafeEdit { operation: "apply_opened_presentation_patch", reason: "an opened-presentation patch resource is stale" }'
    output = (P / 'capture-v3/artifacts/owned-control-output.pptx').read_bytes()
    partial = gates['partial_sink']
    assert partial['status'] == 'verified' and partial['error_kind'] == 'OpcError::IncompleteOutput'
    assert partial['accepted_prefix_bytes'] == len(output) // 2 == 8508572
    assert sha(output[:8508572]) == partial['prefix_sha256']
    zero = gates['zero_sink']
    assert zero['status'] == 'verified' and zero['accepted_prefix_bytes'] == 0
    assert zero['error_kind'] == 'OpcError::ZipError (preservation writer mapping)' and zero['prefix_sha256'] == sha(b'')
    assert gates['input_limit_actual'] == 17017139 and gates['input_limit_maximum'] == 17017138
    runs = load('census-runs.json')
    check(runs['helper'])
    assert runs['status'] == 'pass' and [r['label'] for r in runs['runs']] == ['output', 'noop', 'inverse', 'forward-replay']
    summary = []
    for row in runs['runs']:
        assert row['exit_code'] == 0
        check(row['report'])
        census = json.loads(Path(row['report']['path']).read_text())
        assert census['status'] == 'pass' and not census['violations']
        for d in census['inputs'].values():
            check(d)
        assert census['ordered_filenames']['equal'] and census['archive']['comment_comparison']['equal']
        assert len(census['members']) == 445
        physical = {side: [m['name'] for m in sorted(census['members'], key=lambda m: m[side]['physical']['local_header_offset'])] for side in ('source', 'output')}
        assert physical['source'] == physical['output'], 'physical order changed'
        expected = SELECTED if row['label'] == 'output' else []
        assert sorted(census['changed_members']['names']) == expected
        assert sorted(census['changed_members']['by_component']['decoded_payload']) == expected
        assert len(census['unchanged_members']) == 445 - len(expected)
        for member in census['members']:
            if member['name'] not in expected:
                assert all(c['equal'] for c in member['components'].values())
        summary.append({'label': row['label'], 'changed_members': expected, 'unchanged_members': 445 - len(expected), 'physical_order_equal': True, 'central_order_equal': True, 'archive_comment_equal': True, 'data_descriptors_observed': census['archive']['source']['data_descriptors'], 'zip64_observed': census['archive']['source']['zip64']})
    return {'status': 'qualified_generated_fixture', 'production_change': False, 'performance_claim': False, 'source': report['source'], 'output': report['output'], 'ownership': report['ownership'], 'gates': gates, 'censuses': summary, 'inverse_scope': 'Live and serialized/reopened inverse restore full bytes of this generated source; logical and semantic equality follow from exact equality to the admitted source. This does not certify arbitrary producer archives.', 'coverage_gaps': report['coverage'], 'next_step': 'Add and measure a comparable owning-input benchmark control with explicit input-copy and replay boundaries before considering a production provenance bridge.', 'bindings': [desc(P / name) for name in ['source-v3.json', 'build-v3/command.json', 'capture-v3/command.json', 'quality-v3-v2/receipt.json', 'census-runs.json', 'qualification-plan.json', 'qualification-amendment.json']]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    assert not args.output.exists()
    result = reduce()
    with args.output.open('x') as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write('\n')
    print('0909 generated-fixture owning control qualified')


if __name__ == '__main__':
    main()
