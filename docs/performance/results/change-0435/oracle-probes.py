#!/usr/bin/env python3
"""Exercise independent report guards against retained pilot mutations."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile

ROOT = Path(__file__).resolve().parent


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    spec = importlib.util.spec_from_file_location('oracle0435', ROOT / 'verify-report.py')
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    source = ROOT / 'pilots/before-buffered/v3/allocator-tiny.json'
    catalog_source = source.with_name('allocator-tiny-catalog.json')
    original = json.loads(source.read_text())
    catalog = json.loads(catalog_source.read_text())
    validations = []
    mutations = [
        ('paragraph-count', lambda r: r['results'][0]['source']['odt_paragraphs'].__setitem__('paragraph_count', 63), 'paragraph count'),
        ('semantic-digest', lambda r: r['results'][0]['source']['odt_paragraphs'].__setitem__('semantic_sha256', '0' * 64), 'independent paragraph digest'),
        ('styles-default', lambda r: r['results'][0]['source']['odt_paragraphs'].__setitem__('styles_xml_sha256', '0' * 64), 'pinned Builder defaults'),
        ('content-target-binding', lambda r: r['results'][0]['source']['odt_paragraphs'].__setitem__('content_xml_sha256', '0' * 64), 'target hash'),
        ('source-role', lambda r: r['results'][0]['source']['odt_paragraphs'].__setitem__('role', 'streaming'), 'role or implementation'),
        ('sink-output', lambda r: r['results'][0]['sink'].__setitem__('accepted_bytes', 1), 'sink summary'),
        ('allocator-region-peak', lambda r: r['results'][0]['operation_metrics']['allocation']['region_peak_live_bytes']['values'].__setitem__(0, 0), 'region peak invariant'),
        ('catalog-reference', lambda r: r['corpus_catalog'].__setitem__('catalog_sha256', '0' * 64), 'catalog sidecar'),
    ]
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0435-oracle-probes-') as temporary:
        directory = Path(temporary)
        report_path = directory / 'report.json'
        catalog_path = directory / 'report-catalog.json'
        catalog_path.write_text(json.dumps(catalog))
        report_path.write_text(json.dumps(original))
        oracle.validate_report(report_path, 'allocator', 'tiny', 'before-buffered', samples=3, warmups=1)
        for name, mutate, expected in mutations:
            report = copy.deepcopy(original)
            mutate(report)
            report_path.write_text(json.dumps(report))
            try:
                oracle.validate_report(report_path, 'allocator', 'tiny', 'before-buffered', samples=3, warmups=1)
            except oracle.VerificationError as error:
                message = str(error)
                assert expected in message, (name, message)
                validations.append({'probe': name, 'rejected': True, 'error': message})
            else:
                raise AssertionError(f'{name}: mutated report was accepted')
    output = {'change': 435, 'status': 'pass', 'driver_sha256': sha(Path(__file__)),
              'oracle_sha256': sha(ROOT / 'verify-report.py'),
              'protocol_sha256': sha(oracle.protocol_path()),
              'source_report': str(source.relative_to(ROOT)), 'source_report_sha256': sha(source),
              'source_catalog_sha256': sha(catalog_source), 'temporary_removed': True,
              'control_validated': True, 'probes': validations}
    with (ROOT / 'preparatory-oracle-probes.json').open('x') as stream:
        json.dump(output, stream, indent=2)
        stream.write('\n')
    print('PASS', len(validations), 'independent oracle mutations')


if __name__ == '__main__':
    main()
