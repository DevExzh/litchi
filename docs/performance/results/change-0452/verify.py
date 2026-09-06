#!/usr/bin/env python3
"""Replay source, build, test, lifecycle, regression and fuzz evidence portably."""
import argparse,hashlib,importlib.util,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def module(name):
    s=importlib.util.spec_from_file_location(name,ROOT/(name+'.py'));m=importlib.util.module_from_spec(s);s.loader.exec_module(m);return m
d=module('derive');oracle=module('verify-report')
def sha(raw):return hashlib.sha256(raw).hexdigest()
def member(name):
    p=Path(name);d.require(not p.is_absolute() and '..' not in p.parts,'safe artifact path')
    p=(ROOT/p).resolve();d.require(p.is_relative_to(ROOT.resolve()),'artifact escape');return p
def artifact(r):
    member(r['path']);raw=d.raw(r['path']);d.require(len(raw)==r['bytes'] and sha(raw)==r['sha256'],'artifact '+r['path']);return raw
def manifest(r):
    member(r['path']);raw=d.raw(r['path']);v=json.loads(raw);d.require(sha(raw)==r['sha256'] and len(v)==r['files'],'source manifest');return v
def check(sealed=False,cleanup=False):
    files=d.load('source-files.json')
    d.require(files==['crates/litchi-opc/src/source_backed.rs','crates/litchi-opc/src/lib.rs','crates/litchi-opc/tests/source_part_transfer.rs','crates/litchi-pptx/src/presentation/source_cross_copy.rs','crates/litchi-pptx/tests/source_backed_cross_copy.rs','crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'],'source set')
    builds={name:d.load(name+'-build.json') for name in ['baseline','candidate']}
    base=manifest(builds['baseline']['source_manifest']);final=manifest(builds['candidate']['source_manifest'])
    for n in files:
        for label,m in [('before',base),('after',final)]:d.require(sha(d.raw('candidate/'+label+'-'+Path(n).name+'.txt'))==m[n],'exact source copy '+label)
    common={k:v for k,v in final.items() if k not in files};d.require({k:v for k,v in base.items() if k not in files}==common,'only reviewed sources differ')
    required=['baseline-harness-build','final-strict-r3','final-opc-tests-r3','final-pptx-tests-r3','final-harness-tests-r3','candidate-harness-build-r3','final-doc-r3','final-workspace-check-r3','final-format-r3','final-boundaries-r3','retention-negative-control','metadata-bound-tests-r3','fuzz-lock','fuzz-build','fuzz-smoke','final-fuzz-strict','final-fuzz-format']
    required += ['confirmation-'+str(i) for i in range(4)]
    required += ['pilot-'+str(i) for i in range(8)]+['formal-'+str(i) for i in range(16)]+[f'profile-{i}-{kind}' for i in [1,3] for kind in ['stat','record']]
    for tag in required:d.require((ROOT/'checks'/f'{tag}.json').is_file(),'required '+tag)
    for p in (ROOT/'checks').glob('*.json'):
        r=d.load(str(p.relative_to(ROOT)))
        if 'source_before' not in r:continue
        d.require(r['change']==452 and r['driver_sha256']==sha(d.raw('check.py')),'check driver')
        d.require(r['source_unchanged'] and r['source_before']==r['source_after'],'source changed during command')
        m=manifest(r['source_after']);d.require({k:v for k,v in m.items() if k not in files}==common,'unrelated source difference')
        if p.stem in required and p.stem not in ['baseline-harness-build','retention-negative-control']:d.require(m==final,'final source scope '+p.stem)
        passed=p.stem not in {'retention-negative-control','metadata-bound-tests','metadata-bound-tests-r2'}
        d.require(r['status']==('pass' if passed else 'failed') and (r['exit_code']==0)==passed,'check status '+p.stem)
        log=artifact(r['log']).decode()
        if 'test' in r['argv']:
            rows=re.findall(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',log)
            for i,key in enumerate(['passed_tests','failed_tests','ignored_tests']):d.require(r[key]==sum(int(v[i]) for v in rows),'derived test count')
    for tag,count in [('final-opc-tests-r3',471),('final-pptx-tests-r3',848),('final-harness-tests-r3',381)]:
        r=d.load('checks/'+tag+'.json');d.require(r['passed_tests']==count and r['failed_tests']==0,'test count '+tag)
    negative=d.load('checks/retention-negative-control.json');old=manifest(negative['source_after'])
    d.require(sha(d.raw('draft-history/pre-retention-fix-source_backed.rs.txt'))==old[files[0]],'negative source identity')
    d.require(negative['passed_tests']==0 and negative['failed_tests']==1 and 'left: 0\n right: 4096' in artifact(negative['log']).decode(),'zero-size retention regression reproduced')
    for name,tag in [('baseline','baseline-harness-build'),('candidate','candidate-harness-build-r3')]:
        r=d.load('checks/'+tag+'.json');d.require(builds[name]['source_manifest']==r['source_after'] and builds[name]['revision']==r['revision'],'build custody')
    protocol=d.load('protocol.json');d.require(protocol['status']=='frozen' and protocol['samples']==30 and protocol['warmups']==3 and protocol['reports']==16 and protocol['retained_samples']==480,'frozen protocol')
    for n,h in protocol['bound_files'].items():d.require(sha(d.raw(n))==h,'frozen dependency '+n)
    orders=[{'provider':provider,'build':build,'corpus':corpus} for provider in ['bytes','range'] for build in ['baseline','candidate'] for corpus in ['plain','media-rich']]
    d.require(protocol['order']==[dict(x,repeat='R1') for x in orders]+[dict(x,repeat='R2') for x in reversed(orders)],'balanced process order')
    confirmation=d.load('confirmation-protocol.json')
    d.require(confirmation['status']=='frozen' and confirmation['cpu']==2 and confirmation['samples']==30 and confirmation['warmups']==3,'confirmation protocol')
    d.require(confirmation['order']==[dict(provider='bytes',corpus='media-rich',build=b,repeat=r) for b,r in [('baseline','R3'),('candidate','R3'),('candidate','R4'),('baseline','R4')]],'confirmation ABBA order')
    for n,h in confirmation['bound_files'].items():d.require(sha(d.raw(n))==h,'confirmation frozen dependency')
    for p in list((ROOT/'confirmation').glob('*/receipt.json'))+list((ROOT/'runs').glob('*/receipt.json'))+list((ROOT/'pilots').glob('*/receipt.json'))+list((ROOT/'profiles').glob('*/*/receipt.json')):
        r=d.load(str(p.relative_to(ROOT)));lane=r['lane'];build=builds[lane['build']]
        d.require(r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0,'capture status')
        d.require(r['source_unchanged'] and r['source_before']==r['source_after']==builds['candidate']['source_manifest'],'capture source epoch')
        d.require(r['binary']==build['binary'] and r['build_sha256']==sha(d.raw(lane['build']+'-build.json')),'capture binary binding')
        is_confirmation=p.is_relative_to(ROOT/'confirmation')
        protocol_name='confirmation-protocol.json' if is_confirmation else 'protocol.json'
        capture_name='confirmation-capture.py' if is_confirmation else 'capture.py'
        if is_confirmation:d.require(lane==confirmation['order'][int(p.parent.name)],'confirmation lane order')
        d.require(r['protocol_sha256']==sha(d.raw(protocol_name)) and r['capture_sha256']==sha(d.raw(capture_name)) and r['oracle_sha256']==sha(d.raw('verify-report.py')),'capture drivers')
        artifacts=r['artifacts'];report=json.loads(artifact(artifacts['report']));oracle.check_report(report)
        for a in artifacts.values():artifact(a)
        d.require(report['binary_sha256']==build['binary']['sha256'] and report['binary_bytes']==build['binary']['bytes'] and report['current_exe']==build['binary']['path'] and report['source_revision']==build['revision'],'reported executable identity')
        d.require(report['corpus']==lane['corpus'] and report['provider']==lane['provider'] and report['samples']==(1 if r['pilot'] else 30) and report['warmup']==(0 if r['pilot'] else 3),'report lane')
        for key,value in protocol['corpora'][lane['corpus']].items():d.require(report[key]==value,'exact corpus/output identity')
        expected=[build['binary']['path'],'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples',str(1 if r['pilot'] else 30),'--warmup',str(0 if r['pilot'] else 3),'--source-revision',build['revision'],'--output',r['argv'][r['argv'].index('--output')+1]]
        if lane['provider']=='range':expected+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
        d.require(r['argv'][-len(expected):]==expected,'exact workload arguments')
    d.require(len(list((ROOT/'runs').glob('*/receipt.json')))==16 and len(list((ROOT/'pilots').glob('*/receipt.json')))==8 and len(list((ROOT/'profiles').glob('*/*/receipt.json')))==4,'capture inventory')
    d.require(len(list((ROOT/'confirmation').glob('*/receipt.json')))==4,'confirmation inventory')
    c=module('confirmation-summary');cv=c.derive();d.require(cv==d.load('confirmation-summary.json') and c.render(cv)==d.raw('confirmation-summary.md').decode(),'derived confirmation')
    profiles=module('profile-summary');d.require(profiles.derive()==d.load('profile-summary.json'),'derived profiles')
    measurements=d.derive();d.require(measurements==d.load('measurements.json') and d.render(measurements)==d.raw('measurements.md').decode(),'derived measurements')
    decision=d.load('decision.json');d.require(decision['goal_complete'] is False and decision['native_coverage_promoted'] is False and decision['decision']=='retain PPTX capture reuse','scoped decision')
    review=d.load('regression-review.json');d.require(review['flags']==measurements['review_flags'] and len(review['dispositions'])==len(review['flags']) and all(x['reviewed'] is True and x['reason'] for x in review['dispositions']),'all flags reviewed')
    d.require(review['confirmation_flags']==cv['review_flags'] and len(review['confirmation_dispositions'])==len(cv['review_flags']) and all(x['reviewed'] is True and x['reason'] and x['flag']==flag for x,flag in zip(review['confirmation_dispositions'],cv['review_flags'])),'confirmation reviewed')
    d.require(all(x['flag']==flag for x,flag in zip(review['dispositions'],review['flags'])),'review binding')
    smoke=d.load('checks/fuzz-smoke.json');d.require('Done 1000 runs' in artifact(smoke['log']).decode() and '-runs=1000' in smoke['argv'] and '-seed=452' in smoke['argv'],'fuzz completion')
    fuzz_build=d.load('checks/fuzz-build.json');d.require('RUSTC_BOOTSTRAP=1' in fuzz_build['argv'] and any('-Z sanitizer=address' in a and '-sanitizer-coverage-level=4' in a for a in fuzz_build['argv']),'instrumented fuzz build')
    prepared=d.load('checks/fuzz-prepared.json');d.require(prepared['status']=='pass' and prepared['task']=='/tmp/litchi-goal-0452-opc-fuzz' and prepared['driver_sha256']==sha(d.raw('fuzz-control.py')),'fuzz preparation')
    inputs={r['path']:r for r in prepared['inputs']};d.require(len(inputs)==18 and inputs['fuzz_targets/parse_opc.rs']['sha256']==final[files[-1]],'fuzz source identity')
    d.require(inputs['Cargo.toml']['sha256']==sha(d.raw('checks/fuzz-Cargo.toml.txt')),'fuzz manifest copy')
    seeds=d.load('seed-manifest.json');d.require(len(seeds)==16,'seed count')
    for r in seeds:
        artifact(r);c=inputs['corpus/'+Path(r['path']).name];d.require(c['sha256']==r['sha256'] and c['bytes']==r['bytes'],'seed copy')
    artifacts=d.load('checks/fuzz-artifacts.json');a={r['path']:r for r in artifacts['artifacts']}
    d.require(artifacts['status']=='pass' and artifacts['task']==prepared['task'] and artifacts['files']==len(a) and artifacts['bytes']==sum(r['bytes'] for r in a.values()),'fuzz artifact inventory')
    for n,r in inputs.items():d.require(a[n]==r,'preserved fuzz input')
    d.require(a['Cargo.lock']['sha256']==sha(d.raw('checks/fuzz-Cargo.lock.txt')),'retained generated lock')
    d.require(a['target/x86_64-unknown-linux-gnu/release/parse_opc']['bytes']>0,'fuzz binary identity')
    if cleanup:
        p=d.load('checks/fuzz-cleanup.json');d.require(p['status']=='pass' and p['temporary_directory_absent'] and p['task']==prepared['task'],'fuzz cleanup result')
        d.require(p['files_removed']==artifacts['files'] and p['bytes_removed']==artifacts['bytes'] and p['artifact_manifest_sha256']==sha(d.raw('checks/fuzz-artifacts.json')),'fuzz cleanup custody')
        d.require(d.load('checks/precleanup.json')['status']=='pass','precleanup replay')
        p=d.load('checks/binary-cleanup.json');d.require(p['status']=='pass' and p['temporary_directory_absent'] and p['task']=='/tmp/litchi-goal-0452-pptx-capture','binary cleanup')
        d.require(p['binaries']=={name:b['binary'] for name,b in builds.items()},'cleaned binary identities')
    if sealed:
        inventory={}
        for line in (ROOT/'SHA256SUMS').read_text().splitlines():
            h,n=line.split('  ',1);d.require(n not in inventory and sha(member(n).read_bytes())==h,'seal member');inventory[n]=h
        d.require(set(inventory)=={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'},'seal coverage')
    return {'status':'pass','change':452,'opc_tests':471,'pptx_tests':848,'harness_tests':381,'formal_reports':16,'formal_samples':480,'confirmation_samples':120,'profiles':4,'fuzz_runs':1000,'sealed':sealed,'cleanup':cleanup}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--sealed',action='store_true');p.add_argument('--cleanup',action='store_true');a=p.parse_args();print(json.dumps(check(a.sealed,a.cleanup)))
