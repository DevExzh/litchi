"""Negative raw-vector checks with custody hashes repaired to reach semantic gates."""
import copy
import json
from pathlib import Path
import shutil
import tempfile

import analyze
from capture import HERE, sha


def main():
    cases = ['phase_alignment', 'generic_counter', 'unmanaged_budget', 'repeat_output']
    results = []
    with tempfile.TemporaryDirectory(prefix='litchi-0520-negative-') as directory:
        scratch = Path(directory)
        for path in HERE.iterdir():
            if path.is_file() and not path.name.startswith('profile-'):
                shutil.copyfile(path, scratch / path.name)
        analyze.HERE = scratch
        analyze.analyze()
        filename = 'native-r1-medium.json'
        receipt_name = 'native-r1-medium.receipt.json'
        original = json.loads((scratch / filename).read_text())
        receipt = json.loads((scratch / receipt_name).read_text())
        for case in cases:
            raw = copy.deepcopy(original)
            row = raw['results'][0]
            source = row['source']['xlsx_cell_values']
            if case == 'phase_alignment':
                source['commit_ns'][0] += 1
            elif case == 'generic_counter':
                row['source']['read_calls'] = [206] * 100
            elif case == 'unmanaged_budget':
                source['budget_used_after_handles_drop'] = [1] * 100
            else:
                source['output_sha256'] = ['0' * 64] * 100
                row['output_sha256'] = '0' * 64
            (scratch / filename).write_text(json.dumps(raw))
            changed = copy.deepcopy(receipt)
            changed['artifacts'][filename] = sha(scratch / filename)
            (scratch / receipt_name).write_text(json.dumps(changed))
            try:
                analyze.analyze()
            except AssertionError:
                results.append(dict(case=case, rejected=True))
            else:
                raise AssertionError('Corrupt evidence admitted: ' + case)
        (scratch / filename).write_text(json.dumps(original))
        restored = copy.deepcopy(receipt)
        restored['artifacts'][filename] = sha(scratch / filename)
        (scratch / receipt_name).write_text(json.dumps(restored))
        analyze.analyze()
    report = dict(status='pass', checks=results, valid_control_and_restoration_pass=True,
                  temporary_directory_removed=not scratch.exists())
    (HERE / 'verifier-tests.json').write_text(json.dumps(report, indent=2) + '\n')
    print(json.dumps(report))


if __name__ == '__main__':
    main()
