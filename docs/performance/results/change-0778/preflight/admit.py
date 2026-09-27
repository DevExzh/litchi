"""Admit only independently checked, qualification-bound exporter artifacts."""
from pathlib import Path
import custody as c
import capture
import oracle

if __name__ == '__main__':
    p = c.P
    assert not (p / 'admission.json').exists()
    plan, source, build = capture.load_inputs()
    assert c.census() == source
    export = c.read(p / 'export.json')
    assert export['exit_code'] == 0 and export['binary'] == build['binaries']['export']
    assert c.artifact(Path(export['binary']['path'])) == export['binary']
    for key, filename in [('source_sha256', 'source.json'), ('plan_sha256', 'plan.json'), ('build_sha256', 'build.json')]:
        assert export[key] == c.sha(p / filename)
    result = oracle.check_manifest(p / 'artifacts-0/manifest.json', require_qualification=True)
    oracle_path = p / 'oracle.json'
    assert not oracle_path.exists()
    c.write(oracle_path, result)
    assert result['status'] == 'pass' and result['qualification_binding']['status'] == 'bound'
    admission = {
        'oracle_pass': True, 'oracle_report': c.artifact(oracle_path),
        'oracle_report_sha256': c.sha(oracle_path), 'oracle_runner_sha256': c.sha(p / 'oracle.py'),
        'qualification_identities_sha256': c.sha(p / 'qualification-identities.json'),
        'export': {'receipt': c.artifact(p / 'export.json'), 'binary': export['binary'],
                   'source_sha256': c.sha(p / 'source.json'), 'source_revision': source['revision'],
                   'plan_sha256': c.sha(p / 'plan.json'), 'build_sha256': c.sha(p / 'build.json'),
                   'manifest': export['manifest']},
        'real_fixtures': {row['id']: {k: row[k] for k in ['path', 'bytes', 'sha256']}
                          for row in plan['corpora'] if row['path']},
        'generated': [row for row in result['corpora'] if row['id'].startswith('generated-')],
    }
    c.write(p / 'admission.json', admission)
    capture.admission_guard(plan, source, build)
    print('Seven corpora admitted with independent output and qualification bindings.')
