#!/usr/bin/env python3
"""Capture the fixed 0498 matched serial/batch matrix, one child at a time."""
import csv
import hashlib
import json
from pathlib import Path
import subprocess
import time

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
CACHE = Path('/home/zhuhe/.cache/litchi-goal-0498')
BINARY = CACHE / 'retained/source_backed_batch_perf-after'
OUT = HERE / 'final'
OUT.mkdir(exist_ok=True)
freeze = json.loads((HERE / 'after-freeze.json').read_text())
assert hashlib.sha256(BINARY.read_bytes()).hexdigest() == freeze['binary_sha256']
for item in json.loads((HERE / 'candidate-source.json').read_text())['files']:
    assert hashlib.sha256((ROOT / item['path']).read_bytes()).hexdigest() == item['sha256'], item['path']
for corpus in ['few-large', 'many-small']:
    for source in ['owned', 'file', 'instrumented']:
        reference = None
        for api, workers in [('serial', 1), ('batch', 1), ('batch', 2), ('batch', 4), ('batch', 8)]:
            name = f'{corpus}-{source}-{api}-w{workers}'
            output = OUT / (name + '.csv')
            receipt = OUT / (name + '.json')
            assert not output.exists() and not receipt.exists(), name
            scratch = CACHE / 'final-corpora' / name
            command = ['taskset', '-c', '0-7', '/usr/bin/time', '-v', str(BINARY),
                       '--corpus', corpus, '--source', source, '--api', api,
                       '--capability', 'managed', '--workers', str(workers),
                       '--warmups', '3', '--samples', '30', '--repeats', '2',
                       '--artifact-dir', str(scratch), '--output', str(output)]
            started = time.time()
            stderr = OUT / (name + '.time.stderr')
            stdout = OUT / (name + '.stdout')
            with stdout.open('wb') as out, stderr.open('wb') as err:
                result = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
            record = dict(command=command, exit_code=result.returncode,
                          started_epoch=started, ended_epoch=time.time(),
                          binary_sha256=freeze['binary_sha256'])
            if result.returncode == 0:
                data = list(csv.DictReader(output.open()))
                rows = [r for r in data if r['record']=='sample' and int(r['index']) >= 3]
                assert len(rows)==60, name
                assert len({(r['logical_bytes'],r['digest']) for r in rows})==1, name
                semantic = (rows[0]['corpus_sha256'], rows[0]['logical_bytes'], rows[0]['digest'])
                if reference is None: reference = semantic
                assert semantic == reference, name
                assert all(int(r['budget_memory_after_release']) == 0 and
                           int(r['budget_objects_after_release']) == 0 for r in rows), name
                assert all(int(r['cache_cold_loads']) == (4 if corpus=='few-large' else 64) for r in rows), name
                assert not scratch.exists(), name
                record.update(measured_samples=60, warmup_samples=6, semantic=list(semantic),
                              csv_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                              stderr_sha256=hashlib.sha256(stderr.read_bytes()).hexdigest(),
                              cleanup_verified=True)
            receipt.write_text(json.dumps(record, indent=2)+'\n')
            print(name, 'exit', result.returncode, flush=True)
            assert result.returncode == 0, name
