#!/usr/bin/env python3
"""Verify the preservation patch, failing control, strict oracle and custody."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def main():
    for name,digest in read(P/'constraints.json').items():assert sha(ROOT/name)==digest,name
    spec=importlib.util.spec_from_file_location('custody0710',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
    baseline=read(P/'baseline/source.json');candidate=read(P/'candidate/source.json');assert C.census()==candidate==read(P/'source-final.json')
    assert {n for n in baseline if baseline[n]!=candidate.get(n)}=={'crates/litchi-docx/src/package/codec.rs'}
    prior_source=read(P.parent/'change-0709/source-final.json')
    assert {n for n in set(prior_source)|set(baseline) if prior_source.get(n)!=baseline.get(n)}=={'crates/litchi-docx/tests/custom_properties.rs'}
    before=read(P/'baseline/checks.json');assert len(before)==1 and before[0]['exit_code']==101
    log=P/'baseline/custom-properties.log';assert sha(log)==before[0]['log_sha256']
    assert before[0]['source_sha256']==sha(P/'baseline/source.json')
    assert 'test result: FAILED. 4 passed; 4 failed;' in log.read_text() and 'error[E' not in log.read_text()
    control=read(P/'baseline-result.json');assert sha(log)==control['log_sha256'] and sha(P/'change.patch')==control['patch_sha256']
    for name in control['failed_tests']:assert 'test '+name+' ... FAILED' in log.read_text()
    for folder,count in [('candidate',3),('oracle',3)]:
        rows=read(P/folder/'checks.json');assert len(rows)==count
        for row in rows:
            assert row['exit_code']==0
            if folder=='candidate':assert row['source_sha256']==sha(P/'candidate/source.json')
            path=P/folder/row.get('log',row['name']+'.log');assert sha(path)==row['log_sha256']
    evidence=read(P/'evidence/results.json');assert len(evidence)==6
    for row in evidence:
        assert row['exit_code']==0 and sha(P/'evidence'/(row['name']+'.log'))==row['log_sha256']
        assert row['source_manifest_sha256']==sha(P/'source-final.json')
    reuse=read(P/'oracle-reuse.json')
    assert sha(P.parent/'change-0709/oracle/report.json')==reuse['source_report_sha256']
    for name,digest in reuse['probe_files'].items():assert sha(P/'oracle'/name)==digest
    result=read(P/'oracle/result.json');report=read(P/'oracle/report.json')
    assert result['exit_code']==0 and report['oracle_pass'] is True
    assert report['successful_fixture']['fixture_pass'] and report['refusal_fixture']['fixture_pass']
    prior=read(P.parent/'change-0709/oracle/report.json')
    assert [r['output_sha256'] for r in report['successful_fixture']['routes']]==[r['output_sha256'] for r in prior['successful_fixture']['routes']]
    assert sha(P/'oracle/source.json')==result['probe_manifest_sha256']
    assert sha(P/'oracle/workspace-source.json')==result['workspace_manifest_sha256']
    assert read(P/'oracle/workspace-source.json')==candidate
    for name,digest in result['artifacts'].items():assert sha(P/'oracle'/name)==digest
    for fixture in ['successful_fixture','refusal_fixture']:
        ident=report[fixture]['identity'];assert sha(ROOT/ident['relative_path'])==ident['source_sha256']
    refused=report['refusal_fixture'];routes=refused['routes']+refused['no_edit_routes'];assert len(routes)==4
    for route in routes:
        assert route['checks']['output_reopened'] and route['checks']['source_and_output_text_equal']
        members=route['members'];assert members['unchanged_compressed_payloads_except_main'] and members['unchanged_decoded_payloads_except_main']
        assert not members['main_member_decoded_payload_changed'] and not members['main_member_compressed_payload_changed']
        assert sha(Path(route['artifact_path']))==route['output_sha256']==refused['identity']['source_sha256']
        assert route['output_bytes']==refused['identity']['source_bytes']
    final=read(P/'final-report-gate.json');assert final['exit_code']==0 and sha(P/'final-report-gate.log')==final['log_sha256']
    for name,digest in final['docs'].items():assert sha(ROOT/name)==digest
    cleanup=read(P/'cleanup.json');assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in cleanup['owned_paths'])
    assert result['binary'] in cleanup['binaries']
    verification=read(P/'zip-verification.json');assert verification['status']=='pass' and len(verification['rows'])==6
    for row in verification['rows']:
        assert sha(P/row['artifact'])==row['sha256']
        if row['fixture']=='refusal_fixture':assert row['changed_decoded_members']==[]
    print('PASS: baseline fails, candidate checks pass, unchanged strict oracle passes, exact source/artifacts and cleanup verified')
if __name__=='__main__':main()
