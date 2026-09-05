#!/usr/bin/env python3
"""Replay an isolated export, then require rejection of corrupted evidence."""
from pathlib import Path
import json
import os
import shutil
import subprocess
import tempfile

root = Path(__file__).resolve().parent
repo = root.parents[3]
env = os.environ.copy()
env['PYTHONDONTWRITEBYTECODE'] = '1'
results = []

with tempfile.TemporaryDirectory(prefix='litchi-goal-0417-replay-') as temporary:
    export = Path(temporary)
    bundle = export / 'docs/performance/results/change-0417'
    shutil.copytree(root, bundle, ignore=shutil.ignore_patterns('__pycache__', 'profile'))
    (export / 'tools').mkdir()
    for name in ('summarize_crud_baseline.py', 'perf_compare.py', 'validate_perf_corpus_binding.py'):
        shutil.copy2(repo / 'tools' / name, export / 'tools' / name)

    def check(label, expect_success):
        argv = ['python3', str(bundle / 'verify.py')]
        process = subprocess.run(argv, cwd=export, env=env, capture_output=True, text=True)
        results.append(dict(label=label, argv=argv, expected_success=expect_success,
                            exit_code=process.returncode, stdout=process.stdout, stderr=process.stderr))
        if (process.returncode == 0) != expect_success:
            raise RuntimeError(f'{label}: unexpected verifier result: {process.stdout} {process.stderr}')

    def corrupt(label, relative, mutation):
        path = bundle / relative
        original = path.read_bytes()
        try:
            value = json.loads(original)
            mutation(value)
            path.write_text(json.dumps(value) + '\n')
            check(label, False)
        finally:
            path.write_bytes(original)

    check('isolated-export-positive', True)
    corrupt('missing-run', 'capture.json', lambda value: value['runs'].pop())
    corrupt('reordered-runs', 'capture.json', lambda value: value['runs'].reverse())
    corrupt('binary-sha-mismatch', 'build-identity.json',
            lambda value: value['binaries']['normal'].__setitem__('sha256', '0' * 64))

    def overlap(value):
        value['runs'][1]['started_utc'] = value['runs'][0]['started_utc']

    corrupt('overlapping-processes', 'capture.json', overlap)

    def wrong_cpu(value):
        value['runs'][0]['argv'][2] = '3'

    corrupt('wrong-affinity', 'capture.json', wrong_cpu)
    capture = json.loads((bundle / 'capture.json').read_text())
    first = capture['runs'][0]['report']
    corrupt('raw-percentile-corruption', first,
            lambda value: value['results'][0]['elapsed_ns'].__setitem__('p50', 1))
    corrupt('source-revision-corruption', first,
            lambda value: value['environment'].__setitem__('git_revision', 'wrong'))
    corrupt('invented-zero-allocation', first,
            lambda value: value['results'][0].__setitem__('operation_metrics', {'allocation': 0}))
    corrupt('summary-corruption', 'summary.json',
            lambda value: value['normal'][0]['repeats'][0].__setitem__('p50_ns', 1))

    for label, relative, suffix in (
        ('taxonomy-byte-corruption', 'inputs/CRUD_Scenario_Checklist.md', b'\ncorruption\n'),
        ('index-byte-corruption', 'inputs/original-crud-index.json', b' '),
    ):
        path = bundle / relative
        original = path.read_bytes()
        try:
            path.write_bytes(original + suffix)
            check(label, False)
        finally:
            path.write_bytes(original)

    path = bundle / 'protocol.json'
    original = path.read_text()
    try:
        path.write_text(original.replace('{', '{"change": 417,', 1))
        check('duplicate-json-key', False)
    finally:
        path.write_text(original)
    check('restored-export-positive', True)

(root / 'checks/negative-probes.json').write_text(json.dumps(results, indent=2) + '\n')
print(json.dumps({'positive_replays': 2, 'corruptions_rejected': len(results) - 2,
                  'temporary_export_removed': True}))
