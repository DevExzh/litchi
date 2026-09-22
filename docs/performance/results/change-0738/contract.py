"""Exact prior oracle contract, shared by capture and offline validation."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / 'change-0735'


def sha(p):
    assert p.is_file() and not p.is_symlink(), p
    return hashlib.sha256(p.read_bytes()).hexdigest()


def read(p):
    return json.loads(p.read_text())


def write(p, value):
    p.write_text(json.dumps(value, indent=2) + '\n')


def validate_sample_shape(sample, native, controls):
    assert isinstance(sample, dict), 'sample object'
    keys = {'index', 'output_sha256', 'output_inventory', 'oracle'}
    keys.add('phase_ns' if native else 'allocations')
    if controls:
        keys.add('retained_witness_count')
    assert set(sample) == keys, 'sample schema'
    assert type(sample['index']) is int
    if controls:
        assert type(sample['retained_witness_count']) is int
    if native:
        assert isinstance(sample['phase_ns'], dict), 'native phase object'
        assert set(sample['phase_ns']) == {'whole_ns'}
        n = sample['phase_ns']['whole_ns']
        assert type(n) is int and n > 0, 'native duration'
    else:
        assert isinstance(sample['allocations'], dict), 'allocation object'
        assert set(sample['allocations']) == {'whole'}
        fields = sample['allocations']['whole']
        assert isinstance(fields, dict), 'whole allocation object'
        assert set(fields) == {'allocated_bytes', 'deallocated_bytes', 'allocation_calls',
                               'peak_live_bytes', 'retained_bytes'}
        assert all(type(v) is int and v >= 0 for v in fields.values()), 'allocation counters'


def validate_report(report, row):
    reference = read(OLD / 'oracle.json')[row['case']]['expected']
    variable = {'schema_version', 'timing_claim', 'allocator_instrumented', 'warmups',
                'samples_requested', 'samples'}
    for key, expected in reference.items():
        if key not in variable:
            assert report[key] == expected, ('report contract', key)
    assert report['warmups'] == row['warmups']
    assert report['samples_requested'] == row['samples']
    assert report['schema_version'] == 1
    controls = row['build'] != 'archive'
    extra = {'lifecycle', 'warmup_receipts', 'retained_witness_count'} if controls else set()
    assert set(report) == set(reference) | extra
    if controls:
        assert report['lifecycle'] == row['lifecycle']
        strict = row['lifecycle'] != 'legacy'
        retained = row['lifecycle'] != 'strict-drained'
        assert type(report['retained_witness_count']) is int
        assert report['retained_witness_count'] == (row['samples'] if retained else 0)
        assert len(report['warmup_receipts']) == (row['warmups'] if strict else 0)
        for i, receipt in enumerate(report['warmup_receipts']):
            validate_sample_shape(receipt, row['lane'] == 'native', True)
            assert receipt['index'] == i and receipt['retained_witness_count'] == 0
            for key in ('output_sha256', 'output_inventory', 'oracle'):
                assert receipt[key] == reference['samples'][0][key], ('warmup oracle', i, key)
    native = row['lane'] == 'native'
    assert report['timing_claim'] is native
    assert report['allocator_instrumented'] is not native
    assert len(report['samples']) == row['samples']
    sample_ref = reference['samples'][0]
    for i, sample in enumerate(report['samples']):
        validate_sample_shape(sample, native, controls)
        assert sample['index'] == i
        if controls:
            assert type(sample['retained_witness_count']) is int
            assert sample['retained_witness_count'] == (i+1 if retained else 0)
        for key in ('output_sha256', 'output_inventory', 'oracle'):
            assert sample[key] == sample_ref[key], ('sample oracle', i, key)
        if native:
            assert set(sample['phase_ns']) == {'whole_ns'}
            assert type(sample['phase_ns']['whole_ns']) is int and sample['phase_ns']['whole_ns'] > 0
            assert sample.get('allocations') is None
        else:
            assert sample.get('phase_ns') is None
            fields = sample['allocations']['whole']
            assert set(fields) == {'allocated_bytes', 'deallocated_bytes', 'allocation_calls',
                                   'peak_live_bytes', 'retained_bytes'}
            assert all(type(v) is int and v >= 0 for v in fields.values())
    return report


def command(row):
    build = read(P / (row['build'] + '-build.json'))
    case = next(c for c in read(P / 'cases.json') if c['id'] == row['case'])
    cmd = ['taskset', '-c', str(read(P / 'plan.json')['cpu']), build['binaries'][row['lane']]['path'],
           '--case', case['case'], '--input', case['path'], '--operation', 'format',
           '--samples', str(row['samples']), '--warmups', str(row['warmups'])]
    if row['build'] != 'archive':
        cmd += ['--lifecycle', row['lifecycle']]
    return cmd
