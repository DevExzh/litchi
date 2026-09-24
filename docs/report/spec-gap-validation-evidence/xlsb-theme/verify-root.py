#!/usr/bin/env python3
"""Independent gate receipts, raw statistics, and XML metadata oracle.

Does not import the Rust harness or its matrix verifier. XML metadata coverage
is the named fields queried by the warm profile, not the full Theme schema.
"""
import hashlib
import json
import math
from pathlib import Path
import re
import statistics
import zipfile
from lxml import etree

BASE = Path(__file__).resolve().parent
ROOT = BASE.parents[3]
RAW = BASE / 'performance/raw/final-release'
CASES = ('source_cold', 'eager_cold', 'source_warm', 'eager_warm', 'noop', 'change_inverse')
SEED = 0xcbf29ce484222325


def digest(data):
    value = SEED
    for byte in data:
        value = ((value ^ byte) * 0x100000001b3) & ((1 << 64) - 1)
    return value


def metadata(xml, changed=False):
    element = etree.fromstring(xml, etree.XMLParser(no_network=True, resolve_entities=False))
    ns = {'a': etree.QName(element).namespace}
    colors = element.find('a:themeElements/a:clrScheme', ns)
    fonts = element.find('a:themeElements/a:fontScheme', ns)
    fields = [element.get('name', ''), colors.get('name'), fonts.get('name'),
              fonts.find('a:majorFont/a:latin', ns).get('typeface'),
              fonts.find('a:minorFont/a:latin', ns).get('typeface')]
    color = colors.find('a:accent1', ns)[0]
    if changed:
        fields.append('010203')
    else:
        fields.append(color.get('val'))
        if etree.QName(color).localname == 'sysClr' and color.get('lastClr'):
            fields.append(color.get('lastClr'))
    return digest(''.join(fields).encode())


def allocation(value):
    assert value['invalid'] is False and value['failed'] == 0
    assert value['live_before'] + value['allocated_bytes'] - value['deallocated_bytes'] == value['live_after']
    assert value['peak_before'] == value['live_before']
    assert value['peak_after'] >= max(value['live_before'], value['live_after'])


def hashes(path, allow_head=False):
    for line in path.read_text().splitlines():
        expected, name = line.split(maxsplit=1)
        if allow_head and expected == 'git_head':
            assert name == (BASE / 'gates/git-rev-before.txt').read_text().strip()
        else:
            assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected, name


def main():
    if not __debug__:
        raise RuntimeError('Verification requires Python assertions enabled')
    gates = BASE / 'gates'
    assert (gates / 'source-hashes-before.txt').read_bytes() == (gates / 'source-hashes-after.txt').read_bytes()
    hashes(gates / 'source-hashes-before.txt')
    assert (gates / 'git-rev-before.txt').read_bytes() == (gates / 'git-rev-after.txt').read_bytes()
    assert (gates / 'workspace-Cargo.lock').read_bytes() == (ROOT / 'Cargo.lock').read_bytes()
    assert (RAW / 'source-state-before.txt').read_bytes() == (RAW / 'source-state-after.txt').read_bytes()
    hashes(RAW / 'source-state-before.txt', True)
    assert (RAW / 'binary-sha256-before.txt').read_bytes() == (RAW / 'binary-sha256-after.txt').read_bytes()
    binary_hash, binary_name = (RAW / 'binary-sha256-before.txt').read_text().split(maxsplit=1)
    binary = Path(binary_name.strip())
    if binary.exists():
        assert hashlib.sha256(binary.read_bytes()).hexdigest() == binary_hash
    provenance = dict(line.split(maxsplit=1) for line in (RAW / 'provenance.txt').read_text().splitlines())
    for key, path in {
        'source_state_before_sha256': RAW / 'source-state-before.txt',
        'source_state_after_sha256': RAW / 'source-state-after.txt',
        'harness_source_sha256': ROOT / 'crates/litchi-xlsb/examples/theme_profile.rs',
        'profile_script_sha256': BASE / 'performance/run-profile.sh',
        'control_script_sha256': BASE / 'performance/make-theme-control.py',
        'schema_verifier_sha256': BASE / 'verify-theme-schema.py',
        'verifier_sha256': BASE / 'performance/verify-matrix.py',
        'release_build_log_sha256': RAW / 'release-build.log',
    }.items():
        assert hashlib.sha256(path.read_bytes()).hexdigest() == provenance[key], key
    assert provenance['binary_sha256_before'] == provenance['binary_sha256_after'] == binary_hash
    counts = {}
    for name in ('theme-all-features-all-targets-incremental0.log', 'theme-all-features-doctests-incremental0.log',
                 'theme-pptx-consumer-unit.log', 'theme-pptx-consumer-integration.log'):
        rows = re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored;', (gates / name).read_text())
        assert rows and all(r[0] == 'ok' and r[2] == '0' for r in rows)
        counts[name] = {'passed': sum(int(r[1]) for r in rows), 'ignored': sum(int(r[3]) for r in rows)}
    fixtures = {
        'testVarious': ROOT / 'test-data/poi/test-data/spreadsheet/testVarious.xlsb',
        '62815': ROOT / 'test-data/ooxml/xlsb/62815.xlsb',
        'theme-control-1048576': RAW / 'generated/theme-control-1048576.xlsb',
    }
    expected = {f'{label}-{case}-p{p}.json' for label in fixtures for case in CASES for p in range(1, 4)}
    assert {p.name for p in RAW.glob('*-p*.json')} == expected
    groups = {}
    for label, path in fixtures.items():
        source = path.read_bytes()
        with zipfile.ZipFile(path) as archive:
            theme = archive.read('xl/theme/theme1.xml')
        before_metadata, after_metadata = metadata(theme), metadata(theme, True)
        source_digest, theme_digest = digest(source), digest(theme)
        for case in CASES:
            reports = []
            for process in range(1, 4):
                report = json.loads((RAW / f'{label}-{case}-p{process}.json').read_text())
                assert report['schema'] == 'xlsb-theme-profile-v1' and report['case'] == case
                assert report['warmup'] == 3 and report['sample_count'] == 30
                assert report['input_bytes'] == len(source) and report['input_digest'] == source_digest
                assert report['theme_bytes'] == len(theme) and report['theme_digest'] == theme_digest
                assert report['expected_metadata_digest'] == before_metadata
                for name in ('semantic_ok', 'digest_stable', 'preservation_ok', 'inverse_ok', 'change_ok', 'source_observation_stable'):
                    assert report[name] is True
                samples = report['samples']
                assert len(samples) == 30
                times = sorted(s['elapsed_ns'] for s in samples)
                assert all(type(t) is int and t > 0 for t in times)
                for percentile in (50, 95, 99):
                    assert report[f'p{percentile}_ns'] == times[math.ceil(30 * percentile / 100) - 1]
                assert report['mean_ns'] == sum(times) // len(times)
                for sample in samples:
                    assert sample['metadata_digest'] == (after_metadata if case == 'change_inverse' else before_metadata)
                    assert sample['source_digest'] == theme_digest
                    for name in ('preservation_ok', 'inverse_ok', 'change_ok'):
                        assert sample[name] is True
                    allocation(sample['allocation'])
                    for key in ('calls', 'requested_bytes', 'returned_bytes'):
                        assert sample['reads_before'][key] + sample['reads'][key] == sample['reads_total'][key]
                    if case == 'source_warm':
                        assert all(v == 0 for v in sample['reads'].values())
                if 'setup_allocation' in report:
                    allocation(report['setup_allocation'])
                reports.append(report)
            groups[f'{label}/{case}'] = {
                'median_process_p50_ns': statistics.median(r['p50_ns'] for r in reports),
                'median_allocated_bytes': statistics.median(s['allocation']['allocated_bytes'] for r in reports for s in r['samples']),
                'median_peak_extra_requested_live_bytes': statistics.median(s['allocation']['peak_after'] - s['allocation']['live_before'] for r in reports for s in r['samples']),
            }
    result = {'baseline_commit': (gates / 'git-rev-before.txt').read_text().strip(),
              'release_binary_sha256': binary_hash, 'tests': counts,
              'reports_verified': 54, 'timed_samples_verified': 1620,
              'independent_xml_oracle': 'name, palette/font names, Latin faces, Accent1; exact raw input/theme digests',
              'groups': groups}
    (BASE / 'root-verification.json').write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
    print(json.dumps(result, indent=2, sort_keys=True))


if __name__ == '__main__':
    main()
