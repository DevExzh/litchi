#!/usr/bin/env python3
"""Capture managed DOCX lifecycle controls from a frozen executable."""
import csv, hashlib, json, subprocess, sys, time
from pathlib import Path
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
CACHE = Path('/home/zhuhe/.cache/litchi-goal-0500')
phase = sys.argv[1]
assert phase in ['before', 'after']
freeze = json.loads((HERE / f'{phase}-freeze.json').read_text())
binary = Path(freeze['binary'])
assert hashlib.sha256(binary.read_bytes()).hexdigest() == freeze['binary_sha256']
assert hashlib.sha256((ROOT / 'crates/litchi-docx/examples/managed_paragraph_batch_perf.rs').read_bytes()).hexdigest() == freeze['harness_sha256']
folder = HERE / phase
folder.mkdir(exist_ok=True)
for paragraphs in [128, 512]:
    for count in [1, 8, 32]:
        for source in ['owned', 'file']:
            for mode in (['repeated'] if phase == 'before' else ['repeated', 'batch']):
                name = f'p{paragraphs}-k{count}-{source}-{mode}'
                output = folder / f'{name}.csv'
                receipt = folder / f'{name}.json'
                assert not output.exists() and not receipt.exists(), name
                scratch = CACHE / 'corpora' / f'{phase}-{name}'
                command = ['taskset', '-c', '0-7', '/usr/bin/time', '-v', str(binary),
                           '--paragraphs', str(paragraphs), '--replacements', str(count),
                           '--source', source, '--mode', mode, '--warmups', '3', '--samples', '30',
                           '--repeats', '2', '--artifact-dir', str(scratch), '--output', str(output)]
                started = time.time()
                stderr = folder / f'{name}.time.stderr'
                stdout = folder / f'{name}.stdout'
                with stdout.open('wb') as out, stderr.open('wb') as err:
                    result = subprocess.run(command, cwd=ROOT, stdout=out, stderr=err)
                record = {'command': command, 'exit_code': result.returncode,
                          'started_epoch': started, 'ended_epoch': time.time(),
                          'binary_sha256': freeze['binary_sha256']}
                if result.returncode == 0:
                    rows = list(csv.DictReader(output.open()))
                    measured = [r for r in rows if r['warmup'] == 'false']
                    assert len(rows) == 66 and len(measured) == 60, name
                    assert len({(r['fixture_sha256'], r['output_sha256'], r['output_bytes']) for r in rows}) == 1, name
                    for row in rows:
                        assert all(row[key] == 'true' for key in ['memory_released', 'objects_released', 'semantic_ok', 'raw_untouched_member_payloads_ok', 'output_exact_ok', 'managed_preflight_forward_ok', 'managed_preflight_inverse_ok', 'source_version_unchanged']), name
                        assert int(row['budget_after_memory']) == int(row['budget_after_objects']) == 0, name
                        assert int(row['budget_after_work']) >= int(row['budget_live_work']) >= int(row['budget_before_work']), name
                    assert not scratch.exists(), name
                    record.update(measured_samples=60, warmup_samples=6, cleanup_verified=True,
                                  csv_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                                  stderr_sha256=hashlib.sha256(stderr.read_bytes()).hexdigest())
                receipt.write_text(json.dumps(record, indent=2) + '\n')
                print(phase, name, 'exit', result.returncode, flush=True)
                assert result.returncode == 0, name
