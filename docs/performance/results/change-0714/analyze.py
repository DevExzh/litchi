#!/usr/bin/env python3
"""Verify source-bound phase reports and isolated publication traces."""
import argparse
import importlib.util
import json
from pathlib import Path

P = Path(__file__).resolve().parent

def module(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value

C = module(P/'capture.py', 'capture0714analysis')
L = module(P.parent/'change-0709/analyze.py', 'legacy0709phase0714')
N = module(P.parent/'change-0713/analyze.py', 'normalizer0713phase0714')

def analyze():
    plan, build, source = C.read(P/'plan.json'), C.read(P/'build.json'), C.read(P/'source.json')
    assert C.C.census() == source == C.read(P.parent/'change-0713/source-final.json')
    assert build['exit_code'] == 0 and build['source_sha256'] == C.C.sha(P/'source.json')
    assert build['log_sha256'] == C.C.sha(P/'build.log')
    for path, sha in C.read(P/'constraints.json').items():
        assert C.C.sha(C.C.ROOT/path) == sha
    for path, sha in C.read(P/'capture-freeze.json').items():
        assert C.C.sha(P/path) == sha
    binary = Path(build['binary']['path'])
    if binary.exists():
        assert C.C.sha(binary) == build['binary']['sha256'] and binary.stat().st_size == build['binary']['bytes']
    else:
        assert C.read(P/'cleanup.json')['binaries'] == [build['binary']]
    meta = dict(binary_sha256=build['binary']['sha256'], binary_bytes=build['binary']['bytes'])
    rows, identities, corpus_identities = [], {}, {}
    for lane in ['native', 'trace']:
        for job in C.jobs(plan, lane):
            name = job['name']
            receipt = C.read(P/(name+'.receipt.json'))
            assert receipt['job'] == job and receipt['command'] == C.command(plan, job, build)
            assert receipt['exit_code'] == 0 and receipt['binary'] == build['binary']
            assert receipt['environment'] == dict(LC_ALL='C', LANG='C', TZ='UTC', RUSTFLAGS=None, LD_PRELOAD=None)
            assert receipt['source_before'] == receipt['source_after'] == C.digest(source)
            for key, filename in [('source_manifest_sha256', 'source.json'), ('build_sha256', 'build.json'),
                                  ('plan_sha256', 'plan.json'), ('script_sha256', 'capture.py')]:
                assert receipt[key] == C.C.sha(P/filename)
            assert receipt['fixture_before'] == receipt['fixture_after'] == C.fixture(job['corpus'])
            suffixes = ['.json', '.stdout', '.stderr'] + (['.strace'] if lane == 'trace' else [])
            assert set(receipt['artifacts']) == {name+suffix for suffix in suffixes}
            for artifact, sha in receipt['artifacts'].items():
                assert C.C.sha(P/artifact) == sha
            report = C.read(P/(name+'.json'))
            L.check_report_metadata(report, meta, 'native', job['case'], job['samples'], job['warmup'], name)
            result = report['results'][0]
            values, stats = L.validate_elapsed(result['elapsed_ns'], job['samples'], name)
            ordinary = L.validate_ordinary_save(result, job['corpus'], job['phase'], 'native', job['samples'], name)
            L.validate_operation_metrics(result['operation_metrics'], result['elapsed_ns'], job['samples'], 'native', name)
            key = job['corpus']['id']+'/'+job['phase']
            identity = N.normalized_result(result)
            if key in identities:
                assert identity == identities[key]
            identities[key] = identity
            corpus_key = job['corpus']['id']
            if corpus_key in corpus_identities:
                assert ordinary['corpus'] == corpus_identities[corpus_key]
            corpus_identities[corpus_key] = ordinary['corpus']
            row = dict(name=name, lane=lane, repeat=job['repeat'], corpus=corpus_key,
                       phase=job['phase'], report_sha256=C.C.sha(P/(name+'.json')),
                       receipt_sha256=C.C.sha(P/(name+'.receipt.json')))
            if lane == 'native':
                row['elapsed_ns'] = L.elapsed_stats(values)
            else:
                parser = module(P/'trace_analysis.py', 'trace0714_'+name.replace('-', '_'))
                row['trace'] = parser.analyze_trace(P/(name+'.strace'), job['phase'],
                                                   plan, ordinary['corpus']['published_bytes'])
            rows.append(row)
    assert len(rows) == 24 and len(identities) == 8 and len(corpus_identities) == 2
    repeat_rows = []
    for corpus in ['generated', 'numbered-list']:
        for phase in plan['native']['phases']:
            selected = [row for row in rows if row['lane'] == 'native' and row['corpus'] == corpus and row['phase'] == phase]
            assert len(selected) == 2
            for metric in ['p50', 'mean', 'p95', 'p99']:
                values = [row['elapsed_ns'][metric] for row in selected]
                spread = (max(values)-min(values))*100/min(values)
                repeat_rows.append(dict(corpus=corpus, phase=phase, metric=metric, values=values,
                                        spread_percent=spread, flag=spread > plan['review_spread_percent']))
    return dict(schema_version=1, revision=plan['revision'], source_sha256=C.C.sha(P/'source.json'),
                plan_sha256=C.C.sha(P/'plan.json'), binary=build['binary'], children=rows,
                native_children=16, trace_children=8, native_repeat_review=repeat_rows,
                normalized_parity=True, phase_corpus_parity=True, corpus_evidence=corpus_identities,
                claims=plan['claims'], performance_claim='none; current-source descriptive diagnostic')

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=P/'analysis.json')
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    encoded = (json.dumps(analyze(), indent=2, sort_keys=True)+'\n').encode()
    if args.check:
        assert args.output.read_bytes() == encoded
    else:
        assert not args.output.exists()
        args.output.write_bytes(encoded)
    print('PASS: 16 native reports, 8 traces, exact source/command/fixture custody and parity')

if __name__ == '__main__':
    main()
