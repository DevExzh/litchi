"""Validate retained native evidence and recompute descriptive phase attribution."""
import json
import math
from pathlib import Path
import statistics
import sys

from capture import HERE, sha, source_check

PHASES = ['open_ns', 'plan_ns', 'commit_ns', 'publication_ns']


def stats(values):
    ordered = sorted(values)
    return dict(p50=(ordered[(len(values)-1)//2] + ordered[len(values)//2])//2,
                p95=ordered[math.ceil(len(values)*.95)-1],
                p99=ordered[math.ceil(len(values)*.99)-1],
                mean=statistics.mean(values), min=min(values), max=max(values))


def analyze():
    source_check()
    plan = json.loads((HERE / 'plan.json').read_text())
    build = json.loads((HERE / 'build.json').read_text())
    rows = []
    same_shape = {}
    receipts = []
    for repeat in range(1, plan['native_repeats'] + 1):
        for shape in plan['shapes']:
            name = f'native-r{repeat}-{shape}'
            receipt = json.loads((HERE / (name + '.receipt.json')).read_text())
            assert receipt['exit_code'] == 0
            assert receipt['binary_sha256'] == build['binary_sha256']
            assert receipt['script_sha256'] == sha(HERE / 'capture.py')
            assert receipt['plan_sha256'] == sha(HERE / 'plan.json')
            assert receipt['source_manifest_sha256'] == sha(HERE / 'source-manifest.json')
            command = receipt['command']
            assert command[:3] == ['taskset', '-c', str(plan['cpu'])]
            for option, value in [('--warmup', plan['warmup']), ('--samples', plan['samples']),
                                  ('--case', plan['case']), ('--xlsx-cell-crud-shape', shape)]:
                assert command.count(option) == 1 and command[command.index(option)+1] == str(value)
            assert set(receipt['artifacts']) == {name + suffix for suffix in ['.json', '.stdout', '.stderr']}
            for filename, digest in receipt['artifacts'].items():
                assert sha(HERE / filename) == digest, filename
            receipts.append(receipt)
            raw = json.loads((HERE / (name + '.json')).read_text())
            assert raw['schema_version'] == 1
            assert raw['binary_identity']['binary_sha256'] == build['binary_sha256']
            assert raw['environment']['git_revision'] == plan['revision']
            assert raw['environment']['cpu_affinity'] == str(plan['cpu'])
            assert raw['tool']['profile'] == 'release'
            assert raw['tool']['instrumentation'] == 'none'
            assert raw['configuration']['cases'] == [plan['case']]
            assert raw['configuration']['xlsx_cell_crud_shapes'] == [shape]
            assert raw['configuration']['samples_per_case'] == plan['samples']
            assert raw['configuration']['warmup_iterations_per_case'] == plan['warmup']
            assert len(raw['results']) == 1
            r = raw['results'][0]
            assert r['case'] == plan['case'] and r['corpus']['shape'] == shape
            n = plan['samples']
            e = r['elapsed_ns']
            assert len(e['samples']) == n
            assert sorted(e['sample_order']) == list(range(n))
            assert sorted(e['samples']) == e['samples']
            measured = stats(e['samples'])
            for metric in measured:
                assert math.isclose(e[metric], measured[metric], rel_tol=1e-12), metric
            sd = statistics.stdev(e['samples'])
            assert math.isclose(sd, e['standard_deviation'], rel_tol=1e-12)
            z = 1.959963984540054
            df = n - 1
            assert df > 30
            critical = z + (z**3+z)/(4*df) + (5*z**5+16*z**3+3*z)/(96*df**2) + (3*z**7+19*z**5+17*z**3-15*z)/(384*df**3)
            margin = critical * sd / math.sqrt(n)
            for key, value in [('lower', max(0, measured['mean']-margin)), ('upper', measured['mean']+margin)]:
                assert math.isclose(e['confidence_interval_95'][key], value, rel_tol=1e-12)
            s = r['source']['xlsx_cell_values']
            assert s['implementation'] == 'source-backed'
            assert s['cache_mode'] == 'unmanaged-control' and not s['cache_budget_managed']
            assert s['update_count'] == math.ceil(r['corpus']['entry_count']/100)
            assert s['selected_worksheet_count'] == 4
            assert s['untouched_member_count'] == 12
            assert s.get('partial_sink_verified') is None
            constants = {}
            for key, values in s.items():
                if isinstance(values, list):
                    assert len(values) == n, key
                    if not key.endswith('_ns'):
                        assert all(value == values[0] for value in values), key
                        constants[key] = values[0]
                else:
                    constants[key] = values
            # This runner retains phase vectors in acquisition order; total
            # statistics are sorted and carry the original sample indexes.
            for sorted_index, original_index in enumerate(e['sample_order']):
                assert sum(s[k][original_index] for k in PHASES) == e['samples'][sorted_index]
            sums = [sum(s[k][i] for k in PHASES) for i in range(n)]
            assert sorted(range(n), key=lambda i: (sums[i], i)) == e['sample_order']
            for generic, specific in [('read_calls', 'source_read_calls'), ('read_bytes', 'source_read_bytes'),
                                      ('ordinary_payload_materializations', 'payload_materializations')]:
                assert r['source'][generic] == s[specific]
            for key, values in r['source'].items():
                if isinstance(values, list):
                    assert len(values) == n and all(v == values[0] for v in values), key
            for key in ['payload_memory_limit', 'publication_planning_memory_headroom', 'cache_budget_memory_limit']:
                assert constants[key] is None, key
            for key in ['pre_publication_budget', 'post_publication_budget']:
                for field, value in constants[key].items():
                    expect_null = field.endswith('_limit') or (key == 'post_publication_budget' and field in ['catalog_reserved_objects', 'cache_reserved_objects'])
                    assert value is None if expect_null else value == 0, (key, field)
            assert not any(constants['output_budget_refusal'].values())
            for key in ['cache_failed_loads', 'cache_evictions', 'cache_bypasses',
                        'cache_oversized_bypasses', 'cache_allocation_bypasses',
                        'cache_in_flight_loads', 'cache_budget_memory_used',
                        'cache_budget_reserved_bytes', 'cache_budget_reservation_failures',
                        'budget_used_after_package_drop', 'budget_used_after_handles_drop',
                        'budget_objects_used_after_handles_drop',
                        'unselected_worksheet_read_calls', 'unselected_worksheet_read_bytes']:
                assert constants[key] == 0, key
            assert constants['payload_materializations'] == constants['cache_successful_loads'] == 6
            assert constants['output_sha256'] == r['output_sha256']
            assert r['sink']['largest_write'] <= 65536
            assert r['sink']['write_size_buckets']['bytes_over_65536'] == 0
            assert sum(r['sink']['write_size_buckets'].values()) == r['sink']['write_calls']
            identity = dict(corpus=r['corpus'], sink=r['sink'], constants=constants)
            assert same_shape.setdefault(shape, identity) == identity, shape
            phases = {k: stats(s[k]) for k in PHASES + ['reopen_ns']}
            shares = {k: sum(s[k])/sum(e['samples']) for k in PHASES}
            assert math.isclose(sum(shares.values()), 1)
            rows.append(dict(name=name, shape=shape, repeat=repeat, samples=n,
                             elapsed_ns=measured, phases=phases, aggregate_time_share=shares))
    receipts.sort(key=lambda r: r['start_utc'])
    assert all(a['end_utc'] <= b['start_utc'] for a, b in zip(receipts, receipts[1:]))
    flags = []
    for shape in plan['shapes']:
        a, b = [r for r in rows if r['shape'] == shape]
        for phase in ['elapsed_ns'] + PHASES + ['reopen_ns']:
            av = a['elapsed_ns'] if phase == 'elapsed_ns' else a['phases'][phase]
            bv = b['elapsed_ns'] if phase == 'elapsed_ns' else b['phases'][phase]
            for metric in ['p50', 'p95', 'p99', 'mean']:
                delta = (bv[metric]/av[metric] - 1)*100
                if abs(delta) > 5:
                    flags.append(dict(shape=shape, phase=phase, metric=metric,
                                      repeat1=av[metric], repeat2=bv[metric], change_percent=delta))
    return dict(status='pass', claim='current phase attribution; no speedup claim',
                total_samples=sum(r['samples'] for r in rows), rows=rows,
                repeat_variation_over_five_percent=flags, identities=same_shape,
                uncertainty='Two fresh children per shape; phase percentiles are descriptive. Within-child samples are not independent host repetitions. Existing raw total-time t intervals are retained, not treated as cross-host confidence.',
                unavailable=['operation-local allocations', 'peak RSS', 'physical I/O',
                             'cold cache', 'parallel scaling', 'native Office producer'])


if __name__ == '__main__':
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / 'analysis.json'
    output.write_text(json.dumps(analyze(), indent=2) + '\n')
    print('Native evidence and phase sums verified:', output)
