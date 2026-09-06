#!/usr/bin/env python3
"""Validate retained check receipts and independently rederive the summary."""
import argparse
import importlib.util
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent


def module(name):
    spec = importlib.util.spec_from_file_location('change0434_' + name, ROOT / (name + '.py'))
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result


def validate(stage):
    v = module('verify')
    planned = v.load(ROOT / 'planned-checks.json')
    for key in ('required_pass', 'required_review'):
        for name in planned[key]:
            p = v.bundle_path(name, key)
            receipt = v.load(p)
            expected = 'pass' if key == 'required_pass' else 'failed'
            if receipt.get('status') != expected:
                v.fail(name, 'required check has wrong terminal status')
    if stage != 'precleanup':
        for name in ('precleanup-portable-v2', 'task-cleanup'):
            if v.load(ROOT / 'checks' / (name + '.json')).get('status') != 'pass':
                v.fail(name, 'required cleanup lifecycle check is not passing')
    expected = v.load(ROOT / 'expected-checks.json')
    actual = {}
    for p in sorted(ROOT.rglob('*.json')):
        if p.name in ('compression.json', 'expected-checks.json'):
            continue
        row = v.load(p)
        if not isinstance(row, dict) or 'source_before' not in row:
            continue
        name = p.relative_to(ROOT).with_suffix('').as_posix()
        actual[name] = row.get('status')
        if row.get('status') not in ('pass', 'failed'):
            v.fail(name, 'receipt is not terminal')
        v.check_source_manifest(row['source_before'], name + '.source_before')
        v.check_source_manifest(row['source_after'], name + '.source_after')
        if row['source_before'] != row['source_after']:
            v.fail(name, 'source changed during recorded command')
        if row.get('log'):
            v.artifact_path(p, row['log'], name + '.log')
        if p.parent == ROOT / 'checks':
            if row.get('driver_sha256') != v.sha(ROOT / 'check.py'):
                v.fail(name, 'command custody driver differs')
            if row.get('status') == 'pass' and row.get('exit_code') != 0:
                v.fail(name, 'passing command has nonzero exit')
            if row.get('status') == 'pass' and row.get('passed_tests', 1) <= 0:
                v.fail(name, 'passing test command ran no tests')
    if actual != expected:
        v.fail('expected-checks.json', 'terminal receipt inventory differs')
    for role in ('before', 'after'):
        for p in sorted((ROOT / role / 'pilots').glob('*-receipt.json')):
            row = v.load(p)
            if row.get('status') != 'pass':
                v.fail(str(p), 'pilot is not passing')
            for name, record in row['artifacts'].items():
                v.artifact_path(p, dict(record, path=name), name)
    summary = module('summary')
    if summary.canonical(v.load(ROOT / 'summary.json')) != summary.canonical(summary.derive()):
        v.fail('summary.json', 'retained summary differs from independent derivation')
    return {'status': 'pass', 'terminal_receipts': len(actual), 'summary_rederived': True}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--stage', choices=('precleanup', 'aftercleanup', 'final'), default='final')
    args = parser.parse_args()
    try:
        print(json.dumps(validate(args.stage), sort_keys=True))
    except (OSError, ValueError, KeyError, TypeError, AssertionError) as error:
        print('INVALID: ' + str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
