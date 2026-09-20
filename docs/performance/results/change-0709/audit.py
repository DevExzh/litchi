#!/usr/bin/env python3
"""Replay the corrected-source 0709 baseline and terminal custody gates."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def read(path):return json.loads(path.read_text())
def main():
    for manifest in ['constraints.json','source-final.json','context-identities.json']:
        for name,digest in read(P/manifest).items():assert sha(ROOT/name)==digest,name
    spec=importlib.util.spec_from_file_location('build0709',P/'build.py')
    B=importlib.util.module_from_spec(spec);spec.loader.exec_module(B)
    assert B.census()==read(P/'source-final.json')==read(P/'source-baseline.json')
    prior=read(P/'metadata-preflight/source-baseline.json');current=B.census()
    assert {n for n in set(prior)|set(current) if prior.get(n)!=current.get(n)}=={'tools/perf-baseline/src/ordinary_save.rs'}
    for fixture in read(P/'fixtures.json'):assert sha(ROOT/fixture['path'])==fixture['sha256']
    for name,digest in read(P/'capture-freeze.json').items():assert sha(P/name)==digest,name
    for rows,expected in [(read(P/'evidence/results.json'),6),(read(P/'quality.json'),4)]:
        assert len(rows)==expected
        for row in rows:
            assert row['exit_code']==0,row['name']
            log=P/row['log'] if 'log' in row else P/'evidence'/(row['name']+'.log')
            assert sha(log)==row['log_sha256']
            manifest=P/row['source_manifest'] if 'source_manifest' in row else P/'source-final.json'
            assert sha(manifest)==row['source_manifest_sha256']
    final=read(P/'final-report-gate.json');assert final['exit_code']==0
    assert sha(P/'final-report-gate.log')==final['log_sha256']
    for name,digest in final['docs'].items():assert sha(ROOT/name)==digest,name
    for name,digest in read(P/'profile-freeze.json').items():assert sha(P/name)==digest,name
    transition=read(P/'profile-revision-transition.json')
    assert transition['build_revision']==read(P/'revision.json')['revision']
    assert sha(P/'profile-revision-transition.patch')==transition['patch_sha256']
    assert sha(P/'source-baseline.json')==transition['source_manifest_sha256']
    names=subprocess.check_output(['git','diff','--name-only',transition['build_revision'],transition['profile_checkout_revision']],cwd=ROOT,text=True).splitlines()
    assert names==transition['changed_paths']==['docs/performance/results/change-0709/'+n for n in ['analyze_profiles.py','profile-plan.json','profile.py']]
    assert subprocess.check_output(['git','diff',transition['build_revision'],transition['profile_checkout_revision']],cwd=ROOT)==(P/'profile-revision-transition.patch').read_bytes()
    negative=read(P/'negative-checks.json')
    assert negative['status']=='pass' and negative['analyzer_sha256']==sha(P/'analyze.py')
    assert len(negative['checks'])==7 and all(r['rejected'] for r in negative['checks'])
    oracle=P/'oracle';result=read(oracle/'result.json');report=read(oracle/'report.json')
    assert result['exit_code']==1 and result['expected_preservation_failure'] is True
    assert report['oracle_pass'] is False and report['successful_fixture']['fixture_pass'] is True
    assert report['refusal_fixture']['fixture_pass'] is False
    assert sha(oracle/'source.json')==result['probe_manifest_sha256']
    assert sha(oracle/'workspace-source.json')==result['workspace_manifest_sha256']
    assert read(oracle/'workspace-source.json')==B.census()
    for name,digest in read(oracle/'source.json').items():assert sha(ROOT/name)==digest
    for name,digest in result['artifacts'].items():assert sha(oracle/name)==digest
    checks=read(oracle/'checks.json');assert len(checks)==3
    for row in checks:assert row['exit_code']==0 and sha(oracle/row['log'])==row['log_sha256']
    refused=report['refusal_fixture'];routes=refused['routes']+refused['no_edit_routes']
    assert len(routes)==4 and len({r['output_sha256'] for r in routes})==1
    for route in routes:
        assert route['checks']['output_reopened'] and route['checks']['source_and_output_text_equal']
        assert route['members']['changed_non_main_decoded_members']==['docProps/custom.xml']
        assert route['members']['main_member_decoded_payload_changed'] is False
        diff=route['members']['changed_non_main_member_diffs'][0]
        assert diff['source_decoded_bytes']==632 and diff['output_decoded_bytes']==602
        assert diff['decoded_bytes_equal'] is False
        artifact=Path(route['artifact_path'])
        assert sha(artifact)==route['output_sha256'] and artifact.stat().st_size==route['output_bytes']
    cleanup=read(P/'cleanup.json')
    assert cleanup['owned_paths_absent'] and all(not Path(n).exists() for n in cleanup['owned_paths'])
    binary=result['binary']
    assert binary in cleanup['binaries']
    for moved in read(P/'metadata-correction.json')['binaries']:
        assert {'path':moved['archive_path'],'sha256':moved['sha256'],'bytes':moved['bytes']} in cleanup['binaries']
    first=read(oracle/'semantic-investigation.json')['binary']
    assert {k:first[k] for k in ['path','sha256','bytes']} in cleanup['binaries']
    scripts=[('analyze.py','analysis.json'),('analyze_profiles.py','profile-analysis.json')]
    with tempfile.TemporaryDirectory(prefix='litchi-0709-audit-') as temporary:
        for script,output in scripts:
            target=Path(temporary)/output
            subprocess.run([sys.executable,'-B',str(P/script),'--output',str(target)],cwd=ROOT,check=True)
            assert target.read_bytes()==(P/output).read_bytes(),output
    print('PASS corrected harness only, exact source, constraints, fixtures, frozen capture, quality/evidence gates and byte-identical analysis replay')
if __name__=='__main__':main()
