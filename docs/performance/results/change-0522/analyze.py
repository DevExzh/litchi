"""Reuse the sealed 0521 numerical verifier for the fresh 0522 plan."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / 'change-0521' / 'analyze.py'
spec = importlib.util.spec_from_file_location('xlsx_0521_numerical', HELPER)
assert spec is not None and spec.loader is not None
BASE = importlib.util.module_from_spec(spec)
spec.loader.exec_module(BASE)
BASE.HERE = HERE
BASE.PLAN = HERE / 'plan.json'
BASE.BOOTSTRAP_SEED = 5220522


def __getattr__(name):
    return getattr(BASE, name)


def analyze(stage=None):
    result = BASE.analyze(stage)
    result['scope'] = '0522 XLSX source-backed cell reference/tag proof comparison'
    result['numerical_verifier'] = dict(path=str(HELPER.relative_to(HERE.parent)),
                                      sha256=hashlib.sha256(HELPER.read_bytes()).hexdigest())
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('output', nargs='?', type=Path)
    parser.add_argument('--output', dest='output_option', type=Path)
    parser.add_argument('--stage', choices=['baseline', 'candidate', 'compare'])
    args = parser.parse_args()
    stage = None if args.stage == 'compare' else args.stage
    output = args.output_option or args.output or HERE / (f'analysis-{stage}.json' if stage else 'comparison.json')
    result = analyze(stage)
    output.write_text(json.dumps(result, indent=2) + '\n')
    print('0522 evidence verified:', output)
