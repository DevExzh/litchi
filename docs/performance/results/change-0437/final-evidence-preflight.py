#!/usr/bin/env python3
"""Revalidate all candidate pilots and compile retained evidence drivers."""
import ast
import copy
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def module(name):
    spec = importlib.util.spec_from_file_location(name.replace('-', '_'), ROOT / (name + '.py'))
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


def main():
    module('evidence-preflight').main()
    v = module('verify')
    lifecycle = module('lifecycle')
    oracle = module('verify-report-candidate')
    summary = module('summary')
    decision = module('decision')
    protocol = json.loads((ROOT / 'oracle-protocol.json').read_text())
    copies = json.loads((ROOT / 'after/binary-copies.json').read_text())
    for path in sorted((ROOT / 'pilots').glob('*/initial/*-receipt.json')):
        receipt = json.loads(path.read_text())
        lifecycle.validate_receipt_path_binding(v, path.relative_to(ROOT).as_posix(), receipt, 'pilot')
        changed = copy.deepcopy(receipt)
        changed['shape'] = 'wrong'
        module('evidence-preflight').reject(
            lambda: lifecycle.validate_receipt_path_binding(v, path.relative_to(ROOT).as_posix(), changed, 'pilot'),
            'fields differ')
    for invalid in ('../summary', '/summary', 'nested/summary'):
        module('evidence-preflight').reject(lambda: lifecycle.module_stem(invalid), 'unsafe lifecycle module')
    count = 0
    for role in ('after-buffered', 'after-streaming'):
        for mode in ('normal', 'allocator'):
            for shape in ('tiny', 'medium', 'large'):
                report = ROOT / 'pilots' / role / 'initial' / f'{mode}-{shape}.json'
                path = report.with_name(f'{mode}-{shape}-receipt.json')
                receipt = json.loads(path.read_text())
                lifecycle.validate_terminal_receipt(v, path, receipt)
                lifecycle.validate_driver_and_oracle_bindings(v, path, receipt, protocol, 'pilot')
                lifecycle.validate_report_and_oracle_artifacts(v, path, receipt)
                lifecycle.revalidate_preparatory_report(v, path, receipt, 'pilot')
                verified = oracle.validate_report(report, mode, shape, role, samples=3, warmups=1)
                identity = v.normalized_identity(verified, report)
                row = {'name': report.name, 'verified': verified, 'lane': {'mode': mode}, 'identity': identity}
                elapsed = summary.elapsed_summary(row)
                assert decision.elapsed_p50({'elapsed_ns': elapsed}, 'pilot') == elapsed['p50']
                summary.allocator_summary(row)
                resource_name = f'pilots/{role}/initial/{mode}-{shape}-resource.log'
                resource = v.read_resource(path, dict(receipt['artifacts'][resource_name], path=resource_name), 'pilot resource')
                assert resource['peak_rss_bytes'] > 0
                artifacts = {key: (p, p.read_bytes()) for key, p in {
                    'report': report,
                    'catalog': report.with_name(f'{mode}-{shape}-catalog.json'),
                    'resource_log': report.with_name(f'{mode}-{shape}-resource.log'),
                }.items()}
                v.check_capture_argv(receipt, dict(protocol, samples=3, warmups=1), {'shape': shape}, copies[mode], artifacts, 'candidate pilot argv')
                count += 1
    for path in sorted((ROOT / 'profiles').rglob('receipt.json')):
        receipt = json.loads(path.read_text())
        kind = 'preparatory-profile' if receipt['preparatory'] else 'formal-profile'
        lifecycle.validate_receipt_path_binding(v, path.relative_to(ROOT).as_posix(), receipt, kind)
        changed = copy.deepcopy(receipt)
        changed['kind'] = 'wrong'
        module('evidence-preflight').reject(
            lambda: lifecycle.validate_receipt_path_binding(v, path.relative_to(ROOT).as_posix(), changed, kind),
            'fields differ')
    scripts = sorted(ROOT.glob('*.py'))
    for path in scripts:
        ast.parse(path.read_text(), filename=str(path))
    print(json.dumps({'status': 'pass', 'candidate_pilots': count, 'syntax_checked_drivers': len(scripts)}))


if __name__ == '__main__':
    main()
