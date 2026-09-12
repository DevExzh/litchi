"""Summarize completed, source-bound quality receipts without rerunning checks."""
import argparse
import json
import re
import checks
import run as R


def summarize(stage):
    folder = R.HERE / stage
    expected = {'check-' + name + '.receipt.json' for name, _ in checks.COMMANDS}
    paths = sorted(folder.glob('check-*.receipt.json'))
    assert {p.name for p in paths} == expected
    rows = []
    for path in paths:
        value = json.loads(path.read_text())
        assert value['exit_code'] == 0 and value['execution_stage'] == stage
        assert value['source_manifest_sha256'] == R.sha(folder / 'source-manifest.json')
        count = sum(int(n) for n in re.findall(r'test result: ok\. (\d+) passed;', path.with_name(path.name.replace('.receipt.json', '.stdout')).read_text()))
        rows.append(dict(name=path.name, receipt_sha256=R.sha(path), executed_tests=count))
    result = dict(status='pass', stage=stage, checks=rows,
                  executed_tests=sum(row['executed_tests'] for row in rows))
    R.write(R.HERE / 'quality-summary.json', result)
    print(json.dumps(dict(status='pass',stage=stage,checks=len(rows),executed_tests=result['executed_tests'])))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--stage', choices=['candidate','final'], required=True)
    summarize(parser.parse_args().stage)
