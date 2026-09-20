#!/usr/bin/env python3
"""Reject tampered reports, receipts and sync calls without changing captures."""
import copy
import importlib.util
import json
from pathlib import Path
import re
import tempfile

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('analysis0714negative', P/'analyze.py')
A = importlib.util.module_from_spec(spec)
spec.loader.exec_module(A)

def main():
    destination = P/'negative-checks.json'
    assert not destination.exists()
    reference = (P/'analysis.json').read_bytes()
    assert (json.dumps(A.analyze(), indent=2, sort_keys=True)+'\n').encode() == reference
    protected = {f.name:A.C.C.sha(f) for f in P.iterdir() if f.is_file()}
    checks = []
    cases = [
        ('elapsed-mean', 'native-r1-generated-edit.json', lambda r:r['results'][0]['elapsed_ns'].__setitem__('mean', r['results'][0]['elapsed_ns']['mean']+1)),
        ('command-cpu', 'native-r1-generated-edit.receipt.json', lambda r:r['command'].__setitem__(2, '13')),
        ('decoded-manifest-size', 'native-r1-numbered-list-edit.json', lambda r:r['results'][0]['corpus'].__setitem__('target_payload_bytes', 1)),
    ]
    original = A.C.read
    for name, filename, mutate in cases:
        def changed(path):
            value = original(path)
            if path == P/filename:
                value = copy.deepcopy(value)
                mutate(value)
            return value
        A.C.read = changed
        rejected = False
        try:
            A.analyze()
        except (AssertionError, RuntimeError, ValueError) as error:
            rejected = True
            checks.append(dict(name=name, rejected=True, error=repr(error)))
        finally:
            A.C.read = original
        assert rejected, name
    T = A.module(P/'trace_analysis.py', 'trace0714negative')
    trace = P/'trace-r1-generated-atomic_publish.strace'
    text = trace.read_text()
    plan = A.C.read(P/'plan.json')
    output = A.C.read(P/'trace-r1-generated-atomic_publish.json')['results'][0]['source']['ordinary_save']['corpus']['published_bytes']
    mutations = []
    lines = text.splitlines(keepends=True)
    matches = [i for i, line in enumerate(lines) if re.search(r' fsync\(\d+<[^>]+/\.litchi-[^>]+\.tmp>\) = 0 ', line)]
    assert len(matches) == 5
    failed_file = list(lines)
    failed_file[matches[-1]] = failed_file[matches[-1]].replace(' = 0 ', ' = -1 EIO (Input/output error) ', 1)
    mutations.append(('failed-measured-file-sync', failed_file))
    parents = [i for i, line in enumerate(lines) if ' fsync(' in line and '.tmp>' not in line]
    assert len(parents) == 5
    # Fail an early setup sync: a later successful call on reused fd must not rescue it.
    failed_parent = list(lines)
    failed_parent[parents[0]] = failed_parent[parents[0]].replace(' = 0 ', ' = -1 EIO (Input/output error) ', 1)
    mutations.append(('failed-setup-parent-sync', failed_parent))
    wrong_path = list(lines)
    wrong_path[matches[-1]] = wrong_path[matches[-1]].replace('.tmp>', '.wrong>')
    mutations.append(('wrong-measured-sync-path', wrong_path))
    for name, modified in mutations:
        with tempfile.TemporaryDirectory(prefix='litchi-0714-negative-') as directory:
            fake = Path(directory)/trace.name
            fake.write_text(''.join(modified))
            rejected = False
            try:
                T.analyze_trace(fake, 'atomic_publish', plan, output)
            except (AssertionError, RuntimeError, ValueError) as error:
                rejected = True
                checks.append(dict(name=name, rejected=True, error=repr(error)))
            assert rejected, name
    assert all(A.C.C.sha(P/name) == sha for name, sha in protected.items())
    A.C.write(destination, dict(status='pass', checks=checks, retained_inputs_unchanged=True,
                               exact_positive_replay=True, temporary_trace_removed=True,
                               analysis_sha256=A.C.C.sha(P/'analysis.json'),
                               analyzer_sha256=A.C.C.sha(P/'analyze.py'),
                               trace_parser_sha256=A.C.C.sha(P/'trace_analysis.py')))
    print('PASS: six corruption refusals, exact positive replay, retained inputs unchanged')

if __name__ == '__main__':
    main()
