"""Supplemental replay for baseline quality, helper counts, and cross-format veto."""
import re
import subprocess
import sys
import custody as c
from profile_analysis import artifact


def check():
    baseline=c.read(c.P/'quality-before.json')
    assert c.read(artifact(baseline['source']))==c.read(c.P/'build-before/source.json')
    assert len(baseline['rows'])==6 and all(r['exit_code']==0 for r in baseline['rows'])
    final=c.read(c.P/'quality.json')
    serial=c.read(c.P/'serial-execution.json')
    artifact(serial['driver'])
    assert [r['name'] for r in serial['rows']]==['build-after','attribute-after','cross-build-after','native','cross-native','allocation','profile']
    previous=final['rows'][-1]['ended']
    for row in serial['rows']:
        assert row['exit_code']==0 and previous<=row['started']<=row['ended']
        artifact(row['log']);previous=row['ended']
    assert [r['command'] for r in baseline['rows']]==[r['command'] for r in final['rows']]
    for r in baseline['rows']:artifact(r['log'])
    tests=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',artifact(baseline['rows'][2]['log']).read_text())
    summary={'suites':len(tests),'passed':sum(int(r[0]) for r in tests),'failed':sum(int(r[1]) for r in tests),'ignored':sum(int(r[2]) for r in tests)}
    assert summary=={'suites':500,'passed':12850,'failed':0,'ignored':89}
    fix=c.read(c.P/'candidate-fix.json')
    failed_source=c.read(artifact(fix['failed_source']))
    failed_checks=c.read(artifact(fix['failed_checks']))
    assert len(failed_checks)==2 and failed_checks[0]['exit_code']==0 and failed_checks[1]['exit_code']!=0
    failure_log=artifact(failed_checks[1]['log']).read_text()
    assert 'error[E0603]' in failure_log and 'trait `BytesStartExt` is private' in failure_log
    after_source=c.read(c.P/'quality-2/source.json')
    changed=[n for n in failed_source['files'] if failed_source['files'][n]!=after_source['files'][n]]
    assert changed==[fix['changed_file']]
    archived=c.P/'candidate/compile-failure-1/after/litchi-ole-common-xml_attributes.rs'
    corrected=c.P/'candidate/compile-failure-2/after/litchi-ole-common-xml_attributes.rs'
    assert c.sha(archived)==fix['before_sha256'] and c.sha(corrected)==fix['after_sha256']
    assert archived.read_text().replace('pub(crate) trait BytesStartExt','pub trait BytesStartExt').replace('pub(crate) struct CheckedAttributes','pub struct CheckedAttributes')==corrected.read_text()
    artifact(fix['failed_archive'])
    attempt2=c.read(c.P/'quality-2/checks.json')
    assert len(attempt2)==2 and attempt2[0]['exit_code']==0 and attempt2[1]['exit_code']!=0
    assert 'error[E0618]' in artifact(attempt2[1]['log']).read_text()
    old_manifest=c.read(c.P/'candidate/compile-failure-2/manifest.json')
    final_source=c.read(c.P/'build-after/source.json')
    expected={r['production'] for r in old_manifest['files']}
    assert {n for n in after_source['files'] if after_source['files'][n]!=final_source['files'][n]}==expected
    for row in old_manifest['files']:
        old=c.P/'candidate/compile-failure-2'/row['after']
        new=c.P/'candidate/format-failure-3'/row['after']
        assert c.sha(old)==row['after_sha256']
        assert old.read_text().replace('let tag = tag(&content);\n            assert_eq!(checked(&tag), quick_xml_until_error(&tag), "{content}");','let element = tag(&content);\n            assert_eq!(checked(&element), quick_xml_until_error(&element), "{content}");')==new.read_text()
    attempt3=c.read(c.P/'quality-3/checks.json')
    assert len(attempt3)==1 and attempt3[0]['exit_code']!=0
    artifact(attempt3[0]['log'])
    for row in old_manifest['files']:
        old=c.P/'candidate/format-failure-3'/row['after']
        new=c.P/'candidate'/row['after']
        assert old.read_text().replace('assert_eq!(checked(&element), quick_xml_until_error(&element), \"{content}\");', 'assert_eq!(\n                checked(&element),\n                quick_xml_until_error(&element),\n                \"{content}\"\n            );')==new.read_text()
    for script in ['attribute_analysis.py','cross_analysis.py','root_audit.py','cross_root_audit.py']:
        subprocess.run([sys.executable,'-B',str(c.P/script),'--check'],check=True)
    helper=c.read(c.P/'attribute-analysis.json');cross=c.read(c.P/'cross-analysis.json')
    assert helper['reports']==2 and helper['iterations']==84 and helper['latency_claim'] is False
    assert cross['counts']=={'native_reports':12,'qualification_reports':1,'native_samples':2880,'qualification_samples':8,'analysis_rows':8}
    assert isinstance(cross['decision']['rejected'],bool)
    assert cross['decision']['rejected']==any(r['reject'] for r in cross['rows'])
    return {'baseline_quality':summary,'helper_reports':2,'helper_iterations':84,
            'cross_reports':13,'cross_samples':2888,'cross_rejected':cross['decision']['rejected'],
            'helper_resource_increase':helper['any_resource_increase']}
