"""Root-owned semantic and qualification acceptance; never overwrites evidence."""
import hashlib
import subprocess
import sys
import time

import custody as c

P = c.P
plan = c.read(P / 'plan.json')
mode = sys.argv[1] if len(sys.argv) == 2 else ''
assert mode in ('artifacts', 'qualification')
path = P / ('artifact-admission.json' if mode == 'artifacts' else 'qualification-admission.json')
assert not path.exists(), f'refusing to overwrite {path}'
started = time.time()
if mode == 'artifacts':
    assert not (P / 'qualification').exists()
    complete = P / 'artifacts.complete.json'
    assert c.read(complete)['cases'] == 6
    attempt = 0
    while (P / f'admission-{attempt}').exists():
        attempt += 1
    out = P / f'admission-{attempt}'
    out.mkdir()
    report = out / 'artifact-audit.json'
    log = out / 'audit.log'
    auditor = c.artifact(P / 'artifact_audit.py')
    manifest = c.artifact(P / 'artifacts/manifest.json')
    command = [sys.executable, '-B', str(P / 'artifact_audit.py'), '--artifacts', str(P / 'artifacts'), '--report', str(report)]
    with log.open('w') as stream:
        result = subprocess.run(command, cwd=c.ROOT, stdout=stream, stderr=subprocess.STDOUT)
    receipt = {'schema': 'litchi.performance.0819.admission-attempt.v1', 'command': command, 'started': started, 'ended': time.time(), 'exit_code': result.returncode, 'auditor': auditor, 'manifest': manifest, 'log': c.artifact(log)}
    if report.exists():
        receipt['report'] = c.artifact(report)
    c.write(out / 'receipt.json', receipt)
    assert result.returncode == 0, f'independent artifact audit failed; retained {out}'
    assert c.artifact(P / 'artifact_audit.py') == auditor
    assert c.artifact(P / 'artifacts/manifest.json') == manifest
    audit = c.read(report)
    assert audit['schema'] == 'litchi.performance.0819.artifact-audit.v1'
    assert audit['ok'] and not audit['errors'] and len(audit['cases']) == 6
    assert all(row['ok'] for row in audit['cases'])
    preservation_path = P / 'zip-preservation.json'
    preservation = c.read(preservation_path)
    assert preservation['schema'] == 'litchi.performance.0819.zip-preservation.v1'
    preservation_descriptor = c.artifact(preservation_path)
    preservation_cases = preservation['cases']
    assert isinstance(preservation_cases, list) and len(preservation_cases) == 6
    assert len({row['case_id'] for row in preservation_cases}) == 6
    exported = c.read(P / 'artifacts/manifest.json')['cases']
    exported_by_id = {row['case_id']: row for row in exported}
    assert set(exported_by_id) == {row['case_id'] for row in preservation_cases}
    for row in preservation_cases:
        exported_case = exported_by_id[row['case_id']]
        source_path = P / 'artifacts' / exported_case['source_archive']['path']
        output_path = P / 'artifacts' / exported_case['policy_outputs'][0]['output']['path']
        assert row['source_sha256'] == c.sha(source_path)
        assert row['output_sha256'] == c.sha(output_path)
        assert row['member_order_equal'] is True
        assert row['archive_comment_equal'] is True
    real = {row['format'].lower(): row for row in exported if row['origin'] == 'caller-named-real-file'}
    assert set(real) == {'docx', 'xlsx', 'pptx'}
    selectors = []
    for case in plan['cases']:
        row = real[case['format']]
        assert row['edit_admitted'] and row['edit_outcome'] == 'admitted'
        expected = c.read(P / 'corpus-inputs.json')[case['input']]
        assert row['source_archive_sha256'] == expected['sha256']
        assert row['source_archive_bytes'] == expected['bytes']
        selectors.append(case | {'source_sha256': row['source_archive_sha256'], 'published_sha256': row['published_sha256'], 'edit_outcome': row['edit_outcome']})
    canonical = P / 'artifact-audit.json'
    assert not canonical.exists()
    canonical.write_bytes(report.read_bytes())
    value = {'schema': 'litchi.performance.0819.artifact-admission.v1', 'accepted': True, 'plan_sha256': c.sha(P / 'plan.json'), 'artifact_complete': c.artifact(complete), 'manifest': manifest, 'audit': c.artifact(canonical), 'auditor': auditor, 'zip_preservation': preservation_descriptor, 'selectors': selectors}
else:
    assert not (P / 'native').exists()
    admission = P / 'artifact-admission.json'
    accepted = c.read(admission)
    assert accepted['accepted'] and accepted['plan_sha256'] == c.sha(P / 'plan.json')
    for key in ('artifact_complete', 'manifest', 'audit', 'auditor', 'zip_preservation'):
        assert c.artifact(accepted[key]['path']) == accepted[key]
    complete = P / 'qualification/complete.json'
    summary = c.read(complete)
    assert summary['reports'] == summary['samples'] == 12
    rows = c.read(P / 'qualification/receipts.json')
    assert len(rows) == 12
    expected = {x['case']: x for x in accepted['selectors']}
    assert {x['case'] for x in rows} == set(expected)
    for row in rows:
        assert row['exit_code'] == 0
        assert c.artifact(row['report']['path']) == row['report']
        report = c.read(row['report']['path'])
        assert len(report['results']) == 1
        result = report['results'][0]
        assert result['case'] == row['case']
        assert len(result['elapsed_ns']['samples']) == 1
        ordinary = result['source']['ordinary_save']
        oracle = expected[row['case']]
        corpus = ordinary['corpus']
        assert corpus['source_archive_sha256'] == oracle['source_sha256']
        assert corpus['published_sha256'] == oracle['published_sha256']
        assert corpus['edit_admitted'] and corpus['edit_outcome'] == oracle['edit_outcome']
        assert corpus['repeated_cycles_identical'] and corpus['repeated_saves_identical']
        assert ordinary['publications_identical'] and ordinary['edit_outcomes_identical']
        # Edit-only deliberately produces no timed publication; its untimed
        # corpus reference and edit-outcome digest still bind the accepted edit.
        hashes = ordinary['published_sha256']
        assert hashes == ([] if oracle['phase'] == 'edit' else [oracle['published_sha256']])
        assert ordinary['edit_outcome_sha256'] == [hashlib.sha256(b'admitted').hexdigest()]
    value = {'schema': 'litchi.performance.0819.qualification-admission.v1', 'accepted': True, 'plan_sha256': c.sha(P / 'plan.json'), 'artifact_admission': c.artifact(admission), 'qualification_complete': c.artifact(complete), 'reports': 12, 'samples': 12}
value.update(started=started, ended=time.time())
c.write(path, value)
print('0819', mode, 'admission PASS')
