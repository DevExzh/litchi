#!/usr/bin/env python3
"""Exercise candidate evidence readers against retained before pilots/profiles."""
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


def reject(call, needle):
    try:
        call()
    except ValueError as error:
        assert needle in str(error), str(error)
        return
    raise AssertionError('mutation was accepted: ' + needle)


def main():
    v = module('verify')
    oracle = module('verify-report-candidate')
    lifecycle = module('lifecycle')
    summary = module('summary')
    hypothesis = module('before-hypothesis')
    protocol = json.loads((ROOT / 'protocol-draft.json').read_text())
    copies = json.loads((ROOT / 'before/binary-copies.json').read_text())
    assert hypothesis.derive() == json.loads((ROOT / 'before-hypothesis.json').read_text())
    count = 0
    for mode in ('normal', 'allocator'):
        for shape in ('tiny', 'medium', 'large'):
            report = ROOT / 'pilots/before-buffered/initial' / f'{mode}-{shape}.json'
            receipt_path = report.with_name(f'{mode}-{shape}-receipt.json')
            receipt = json.loads(receipt_path.read_text())
            lifecycle.validate_terminal_receipt(v, receipt_path, receipt)
            lifecycle.validate_driver_and_oracle_bindings(v, receipt_path, receipt, protocol, 'pilot')
            lifecycle.validate_report_and_oracle_artifacts(v, receipt_path, receipt)
            lifecycle.revalidate_preparatory_report(v, receipt_path, receipt, 'pilot')
            verified = oracle.validate_report(report, mode, shape, 'before-buffered', samples=3, warmups=1)
            identity = v.normalized_identity(verified, report)
            row = {'name': report.name, 'verified': verified, 'lane': {'mode': mode}, 'identity': identity}
            summary.elapsed_summary(row)
            summary.allocator_summary(row)
            artifacts = {key: (path, path.read_bytes()) for key, path in {
                'report': report,
                'catalog': report.with_name(f'{mode}-{shape}-catalog.json'),
                'resource_log': report.with_name(f'{mode}-{shape}-resource.log'),
            }.items()}
            pilot_protocol = dict(protocol, samples=3, warmups=1)
            v.check_capture_argv(receipt, pilot_protocol, {'shape': shape}, copies[mode], artifacts, 'pilot argv')
            changed = copy.deepcopy(receipt)
            changed['argv'][2] = '3'
            reject(lambda: v.check_capture_argv(changed, pilot_protocol, {'shape': shape}, copies[mode], artifacts, 'pilot argv'), 'CPU affinity')
            count += 1
    for kind in ('stat', 'record'):
        receipt_path = ROOT / 'profiles/preparatory-before-buffered/initial' / kind / 'receipt.json'
        receipt = json.loads(receipt_path.read_text())
        lifecycle.validate_terminal_receipt(v, receipt_path, receipt)
        lifecycle.validate_driver_and_oracle_bindings(v, receipt_path, receipt, protocol, 'profile')
        lifecycle.validate_report_and_oracle_artifacts(v, receipt_path, receipt)
        lifecycle.revalidate_preparatory_report(v, receipt_path, receipt, 'profile')
        artifacts = {name: v.artifact_data(receipt_path, record, name) for name, record in receipt['artifacts'].items()}
        v.check_profile_argv(receipt, protocol, copies['normal'], artifacts, kind, 'profile argv')
        if kind == 'record':
            changed = copy.deepcopy(receipt)
            changed['record_event'] = 'instructions:u'
            reject(lambda: v.check_profile_argv(changed, protocol, copies['normal'], artifacts, kind, 'profile argv'), 'sampling/call-graph')
    print(json.dumps({'status': 'pass', 'pilots': count, 'preparatory_profiles': 2, 'mutation_checks': 7, 'hypothesis_unchanged': True}))


if __name__ == '__main__':
    main()
