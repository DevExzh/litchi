"""Replay the descriptive native capture identities and process summaries."""
import hashlib
import json
import math
from pathlib import Path
import statistics

P = Path(__file__).resolve().parent


def read(name):
    return json.loads((P / name).read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def analyze():
    build = read('build.json')
    source = read('source.json')
    manifest = read('captures/manifest.json')
    assert len(manifest) == 3
    corpus_names = {f'{shape}-{payload}' for shape in ('tiny', 'many-small', 'few-large', 'wide-root')
                    for payload in ('compressible', 'incompressible')}
    keys = {(case, corpus) for case in ('opc_mutated_save', 'opc_noop_save') for corpus in corpus_names}
    groups = {key: [] for key in keys}
    previous_end = 0
    projection = read('descriptive-projection.json')
    for index, receipt in enumerate(manifest):
        report_path = P / 'captures' / f'{index:02d}.json'
        assert receipt['repeat'] == index and receipt['exit'] == 0
        assert previous_end <= receipt['monotonic_start'] < receipt['monotonic_end']
        previous_end = receipt['monotonic_end']
        assert receipt['build_sha256'] == sha(P / 'build.json')
        assert receipt['command'] == ['taskset', '-c', '12', build['binary'], '--warmup', '2',
                                      '--samples', '9', '--case', 'opc_mutated_save,opc_noop_save',
                                      '--json', str(report_path)]
        expected_files = {str(f.relative_to(P)) for f in
                          (report_path, report_path.with_suffix('.stdout'), report_path.with_suffix('.stderr'))}
        assert set(receipt['files']) == expected_files
        for name, digest in receipt['files'].items():
            assert sha(P / name) == digest, name
        report = json.loads(report_path.read_text())
        assert report['schema_version'] == 1
        assert report['environment']['cpu_affinity'] == '12'
        assert report['environment']['git_revision'] == source['head']
        identity = report['binary_identity']
        assert set(identity) == {'path', 'binary_sha256', 'binary_bytes', 'mode_bits', 'executable', 'profile'}
        for key, expected in (('path', build['binary']), ('binary_sha256', build['binary_sha256']),
                              ('binary_bytes', build['binary_bytes']), ('executable', True), ('profile', 'release')):
            assert identity[key] == expected
        assert type(identity['mode_bits']) is int and 0 <= identity['mode_bits'] <= 0o7777
        assert identity['mode_bits'] & 0o111
        configuration = report['configuration']
        assert configuration['samples_per_case'] == 9 and configuration['warmup_iterations_per_case'] == 2
        assert configuration['cases'] == ['opc_mutated_save', 'opc_noop_save']
        assert configuration['corpus_shapes'] == ['tiny', 'many-small', 'few-large', 'wide-root']
        assert configuration['payload_kinds'] == ['compressible', 'incompressible']
        observed = set()
        for result in report['results']:
            key = result['case'], result['corpus']['name']
            assert key in keys and key not in observed
            observed.add(key)
            stable = {name: result.get(name) for name in ('corpus', 'sink', 'output_sha256')}
            assert stable == projection['/'.join(key)]
            elapsed = result['elapsed_ns']
            samples = elapsed['samples']
            assert len(samples) == 9 and samples == sorted(samples)
            assert all(type(value) is int and value > 0 for value in samples)
            assert sorted(elapsed['sample_order']) == list(range(9))
            median = statistics.median(samples)
            p95 = samples[math.ceil(.95 * len(samples)) - 1]
            mean = statistics.mean(samples)
            assert elapsed['p50'] == median and elapsed['p95'] == p95
            assert math.isclose(elapsed['mean'], mean, rel_tol=1e-12)
            groups[key].append({'repeat': index, 'p50_ns': median, 'p95_ns': p95, 'mean_ns': mean})
        assert observed == keys
    output = []
    for key in sorted(keys):
        processes = groups[key]
        medians = [row['p50_ns'] for row in processes]
        spread = max(medians) / min(medians) - 1
        output.append({'case': key[0], 'corpus': key[1], 'processes': processes,
                       'median_process_p50_ns': statistics.median(medians),
                       'process_p50_spread_percent': spread * 100,
                       'spread_above_five_percent': spread > .05})
    return {'status': 'pass', 'source_head': source['head'], 'processes': 3,
            'case_results': 48, 'samples': 432, 'results': output,
            'scope': 'Current-source descriptive ordinary build. No old/new ratio, attribution, or optimization acceptance. '
                     'The selector retains its existing same-writer expected-byte oracle; published hashes and independent '
                     'raw-member preservation proofs are not exposed by this harness revision.'}


if __name__ == '__main__':
    value = analyze()
    (P / 'analysis.json').write_text(json.dumps(value, indent=2) + '\n')
    print('PASS 3 serial processes / 48 case results / 432 samples')
