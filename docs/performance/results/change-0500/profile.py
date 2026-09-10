#!/usr/bin/env python3
"""Record separate whole-child counters; these are not timed-phase attribution."""
import hashlib, json, subprocess, sys
from pathlib import Path
HERE = Path(__file__).resolve().parent
CACHE = Path('/home/zhuhe/.cache/litchi-goal-0500')
phase = sys.argv[1]
assert phase in ['before', 'after']
freeze = json.loads((HERE / f'{phase}-freeze.json').read_text())
binary = Path(freeze['binary'])
assert hashlib.sha256(binary.read_bytes()).hexdigest() == freeze['binary_sha256']
outdir = HERE / 'profiles' / phase
outdir.mkdir(parents=True, exist_ok=True)
for count in [1, 32]:
    for mode in (['repeated'] if phase == 'before' else ['repeated', 'batch']):
        name = f'p512-k{count}-owned-{mode}'
        receipt = outdir / f'{name}.json'
        assert not receipt.exists()
        scratch = CACHE / 'profile-corpora' / f'{phase}-{name}'
        command = ['taskset', '-c', '0-7', 'perf', 'stat', '-x,', '-o', str(outdir / f'{name}.perf.csv'),
                   '-e', 'cycles,instructions,branches,branch-misses,cache-misses,context-switches,cpu-migrations,page-faults',
                   '--', str(binary), '--paragraphs', '512', '--replacements', str(count), '--source', 'owned',
                   '--mode', mode, '--warmups', '3', '--samples', '30', '--repeats', '2',
                   '--artifact-dir', str(scratch), '--output', str(outdir / f'{name}.samples.csv')]
        with (outdir / f'{name}.stdout').open('wb') as stdout, (outdir / f'{name}.stderr').open('wb') as stderr:
            result = subprocess.run(command, stdout=stdout, stderr=stderr)
        record = {'command': command, 'exit_code': result.returncode,
                  'binary_sha256': freeze['binary_sha256'], 'cleanup_verified': not scratch.exists(),
                  'scope': 'whole-child including fixture setup and independent verification'}
        receipt.write_text(json.dumps(record, indent=2) + '\n')
        assert result.returncode == 0 and not scratch.exists(), name
        print(phase, name, 'passed', flush=True)
