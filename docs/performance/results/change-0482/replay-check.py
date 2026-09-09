#!/usr/bin/env python3
"""Verify fresh copies of the actual bundle and retain rejection evidence."""
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

from common import ROOT, TEMP, meta, now, read, write


def mutate(bundle, case):
    if case == 'missing-capture':
        next((bundle / 'captures').glob('*.report.json')).unlink()
    elif case == 'tampered-summary':
        path = bundle / 'summary.json'
        value = read(path)
        value['unrecorded_result'] = True
        path.write_text(json.dumps(value))
    elif case == 'missing-required-gate':
        (bundle / 'validation/opc-tests-accepted.json').unlink()
    elif case == 'tampered-fuzz-record':
        path = bundle / 'fuzz/accepted/smoke.json'
        value = read(path)
        value['exit_code'] = 1
        path.write_text(json.dumps(value))


def main():
    TEMP.mkdir(parents=True, exist_ok=True)
    output = ROOT / 'evidence-validation'
    output.mkdir(exist_ok=True)
    cases = ('portable-baseline', 'missing-capture', 'tampered-summary',
             'missing-required-gate', 'tampered-fuzz-record')
    results = []
    for case in cases:
        with tempfile.TemporaryDirectory(prefix='actual-replay-', dir=TEMP) as scratch:
            bundle = Path(scratch) / 'checkout/docs/performance/results/change-0482'
            shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns('__pycache__'))
            mutate(bundle, case)
            argv = [sys.executable, '-B', str(bundle / 'verify.py')]
            started = now()
            stdout = output / f'{case}.stdout'
            stderr = output / f'{case}.stderr'
            with stdout.open('x') as out, stderr.open('x') as err:
                result = subprocess.run(argv, cwd=bundle, stdout=out, stderr=err)
            accepted = result.returncode == 0
            expected = case == 'portable-baseline'
            record = dict(case=case, argv=argv, cwd=str(bundle), started_utc=started,
                          finished_utc=now(), exit_code=result.returncode,
                          expected_acceptance=expected, accepted=accepted,
                          stdout=meta(stdout), stderr=meta(stderr))
            write(output / f'{case}.json', record)
            results.append(record)
            print(case, result.returncode, flush=True)
            assert accepted == expected, f'{case}: unexpected verifier outcome'
    write(output / 'replay-check.json', dict(
        helper=meta(Path(__file__)), finished_utc=now(), cases=results, status='pass'))


if __name__ == '__main__':
    main()
