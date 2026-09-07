#!/usr/bin/env python3
"""Replay the exact matched capture, release gates, derivation and cleanup custody."""
import argparse,contextlib,hashlib,importlib.util,io,json,re,shutil,subprocess,sys,tempfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
BUNDLE=ROOT.parent
def load(p):return json.loads(p.read_text())
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def module(name,path):
    spec=importlib.util.spec_from_file_location(name,path);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m
def artifact(base,row):
    p=base/row['path'];assert p.resolve().is_relative_to(base.resolve()) and p.is_file() and not p.is_symlink(),row
    assert sha(p)==row['sha256'],p
    if 'bytes' in row:assert p.stat().st_size==row['bytes'],p
    return p
def manifest(base,row):
    p=artifact(base,row);m=load(p);assert len(m)==row['files'];return m
def gate(base,name,expected_source=None,argv=None):
    r=load(base/'checks'/(name+'.json'));assert r['status']=='pass' and r['exit_code']==0 and r['source_unchanged'],name
    assert r['change']==456 and r['revision']==load(ROOT/'final-bindings.json')['revision']
    assert r['source_scope']=='workspace and standalone tools'
    env={'RUSTUP_TOOLCHAIN':'1.98.1','CARGO_BUILD_JOBS':'4','CARGO_INCREMENTAL':'0','PYTHONDONTWRITEBYTECODE':'1'}
    if 'doc' in r['argv']:env['RUSTDOCFLAGS']='-D warnings'
    assert r['environment']==env
    assert r['source_before']==r['source_after'];manifest(base,r['source_after']);artifact(base,r['log'])
    assert sha(base/'check.py')==r['driver_sha256']
    if expected_source:assert r['source_after']==expected_source,name
    if argv:assert r['argv']==argv,name
    if 'test' in r['argv']:
        rows=re.findall(rb'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',artifact(base,r['log']).read_bytes())
        assert [r[k] for k in ['passed_tests','failed_tests','ignored_tests']]==[sum(int(row[i]) for row in rows) for i in range(3)]
        assert r['passed_tests']>0 and r['failed_tests']==0
    return r
def capture_identity(base,relative,receipt,build,index):
    lane=receipt['lane_definition'];assert receipt['lane']==index
    assert artifact(base,receipt['build_manifest'])==base/(lane['build']+'-build.json')
    assert receipt['binary']==build['binaries'][lane['instrumentation']]
    original=Path(receipt['cwd'])/relative;directory=original/'runs'/str(index)
    expected=['taskset','-c','2','/usr/bin/time','-v','-o',str(directory/'resource.log'),receipt['binary']['path'],'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples','30','--warmup','3','--source-revision',build['revision'],'--output',str(directory/'report.json')]
    if lane['provider']=='range':expected+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
    assert receipt['argv']==expected
    assert receipt['oracle_argv']==['/usr/bin/python3','-B',str(original/'verify-report.py'),str(directory/'report.json')]
    report=load(base/'runs'/str(index)/'report.json')
    assert report['binary_sha256']==receipt['binary']['sha256'] and report['binary_bytes']==receipt['binary']['bytes']
    assert report['source_revision']==build['revision'] and report['samples']==30 and report['warmup']==3 and report['checked_iteration_count']==33
    assert report['corpus']==lane['corpus'] and report['provider']==lane['provider']
    assert report['instrumentation']==('none' if lane['instrumentation']=='normal' else 'system_allocator_operation_scoped')
def check_sink_consistency(report):
    bounds={'bytes_0':(0,0),'bytes_1_to_512':(1,512),'bytes_513_to_4096':(513,4096),'bytes_4097_to_16384':(4097,16384),'bytes_16385_to_65536':(16385,65536),'bytes_over_65536':(65537,2**64-1)}
    for row in report['samples_raw']:
        sink=row['publication_sink'];occupied=[bounds[k] for k,n in sink['write_size_buckets'].items() if n]
        low,high=max(occupied)
        assert low<=sink['largest_write']<=high and sink['largest_write']<=sink['accepted_bytes'],'sink largest write contradicts histogram'
        assert sink['accepted_bytes']<=sink['write_calls']*sink['largest_write'],'sink maximum cannot account for accepted bytes'
def check_markdown(base):
    with tempfile.TemporaryDirectory(prefix='litchi-0456-render-') as directory:
        target=Path(directory);shutil.copyfile(base/'protocol.json',target/'protocol.json');shutil.copytree(base/'runs',target/'runs')
        renderer=module('markdown_renderer',base/'derive.py');renderer.ROOT=target
        with contextlib.redirect_stdout(io.StringIO()):renderer.main()
        for name in ['measurements.json','measurements.md']:assert (base/name).read_bytes()==(target/name).read_bytes(),name+' differs from regenerated output'
def verify(precleanup=False,portable=False):
    assert not (precleanup and portable)
    assert sha(ROOT/'final-bindings.json')=='1c3cc7e39bce8a41dde84de8e802f71b2d2227d558a86f0a6ae5f0a3473d2b63'
    bindings=load(ROOT/'final-bindings.json')
    for name,digest in bindings['files'].items():assert sha(ROOT/name)==digest,name
    for name,digest in bindings['bundle_files'].items():assert sha(BUNDLE/name)==digest,name
    assert set(bindings)=={'revision','files','bundle_files','changed_sources','source_snapshots'}
    protocol=load(ROOT/'protocol.json');assert protocol['change']==456 and protocol['status']=='frozen'
    assert protocol['samples']==30 and protocol['warmups']==3 and len(protocol['lanes'])==24
    for name,digest in protocol['bound_files'].items():assert sha(ROOT/name)==digest,name
    canonical=[]
    for round,kind in enumerate(['baseline','candidate','candidate','baseline']):
        order=[('bytes','plain'),('bytes','media-rich'),('range','plain'),('range','media-rich')]
        if round>=2:order.reverse()
        for provider,corpus in order:canonical.append(dict(build=kind,provider=provider,corpus=corpus,repeat='R1' if round<2 else 'R2',instrumentation='normal'))
    for round,kind in enumerate(['baseline','candidate','candidate','baseline']):
        for corpus in (['plain','media-rich'] if round<2 else ['media-rich','plain']):canonical.append(dict(build=kind,provider='bytes',corpus=corpus,repeat='R1' if round<2 else 'R2',instrumentation='allocator'))
    assert protocol['lanes']==canonical and protocol['cpu']==2 and protocol['workers']==1 and protocol['review_percent']==5
    for key,value in {'max_range_bytes':65536,'delay_us':200,'transfer_bytes_per_second':26214400,'transfer_delay_policy':'separate-sleeps'}.items():assert protocol['range'][key]==value
    builds={name:load(ROOT/(name+'-build.json')) for name in ['baseline','candidate']}
    manifests={name:manifest(ROOT,b['source_manifest']) for name,b in builds.items()}
    assert manifests['baseline'].keys()==manifests['candidate'].keys()
    changed={p for p,h in manifests['baseline'].items() if manifests['candidate'][p]!=h}
    assert changed==set(bindings['changed_sources'])=={'crates/soapberry-zip/src/preserve.rs','crates/soapberry-zip/src/writer.rs','crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs'}
    assert {(row['build'],row['source']) for row in bindings['source_snapshots']}=={(kind,source) for kind in ['baseline','candidate'] for source in changed}
    assert len(bindings['source_snapshots'])==6
    for kind,b in builds.items():
        assert b['revision']==bindings['revision'] and b['build_receipt']=='checks/'+kind+'-build.json'
        argv=['cargo','build','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc']
        if kind=='candidate':argv+=['--bin','pptx_external_cross_copy']
        gate(ROOT,kind+'-build',b['source_manifest'],argv)
        for binary in b['binaries'].values():
            p=Path(binary['path'])
            if precleanup:assert p.is_file() and p.stat().st_size==binary['bytes'] and sha(p)==binary['sha256']
    for row in bindings['source_snapshots']:
        assert sha(ROOT/row['snapshot'])==manifests[row['build']][row['source']]
    failed=load(ROOT/'checks/zip-initial.json')
    assert failed['status']=='failed' and failed['exit_code']==101 and failed['source_unchanged']
    assert failed['source_before']==failed['source_after'] and failed['driver_sha256']==sha(ROOT/'check.py')
    assert failed['argv']==module('initial_commands',ROOT/'run-checks.py').COMMANDS['zip-initial-r1']
    failed_sources=manifest(ROOT,failed['source_after']);artifact(ROOT,failed['log'])
    assert failed_sources.keys()==manifests['candidate'].keys()
    assert {p for p,h in failed_sources.items() if manifests['candidate'][p]!=h}=={'crates/soapberry-zip/src/writer.rs'}
    assert sha(ROOT/'failed-initial-writer.rs.txt')==failed_sources['crates/soapberry-zip/src/writer.rs']
    assert (ROOT/'failed-initial-writer.rs.txt').read_bytes().replace(b'self.local_header.fixed.write(writer)?;',b'self.local_header.fixed.write(&mut *writer)?;')==(ROOT/'candidate-writer.rs.txt').read_bytes()
    oracle=module('oracle',ROOT/'verify-report.py')
    assert {p.name for p in (ROOT/'runs').iterdir()}=={str(i) for i in range(24)}
    for i,lane in enumerate(protocol['lanes']):
        d=ROOT/'runs'/str(i);r=load(d/'receipt.json');report=load(d/'report.json');build=builds[lane['build']]
        assert r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0 and r['lane_definition']==lane
        expected_source=builds['baseline' if i in [0,1,2,3,16,17] else 'candidate']['source_manifest']
        assert r['source_before']==r['source_after']==expected_source and r['source_unchanged']
        assert r['protocol_sha256']==sha(ROOT/'protocol.json') and r['capture_sha256']==sha(ROOT/'capture.py')
        assert set(r['artifacts'])=={'report.json','resource.log','workload.log','oracle.log'}
        for name,item in r['artifacts'].items():
            assert item['path']=='runs/'+str(i)+'/'+name
            artifact(ROOT,item)
        capture_identity(ROOT,'docs/performance/results/change-0456/formal',r,build,i)
        for key,value in protocol['corpora'][lane['corpus']].items():assert report[key]==value
        oracle.check_report(report);check_sink_consistency(report)
    commands=module('commands',ROOT/'run-checks.py').COMMANDS;tests=0
    for name,argv in commands.items():r=gate(ROOT,name,builds['candidate']['source_manifest'],argv);tests+=r.get('passed_tests',0)
    assert tests==2654 and len(commands)==11
    for name,argv in module('fuzz_commands',ROOT/'run-fuzz.py').COMMANDS.items():gate(ROOT,name,builds['candidate']['source_manifest'],argv)
    fuzz=load(BUNDLE/'fuzz-inputs/manifest.json')
    assert len(fuzz)==8 and load(BUNDLE/'fuzz-inputs/seed-manifest.json')==fuzz[:6]
    assert {row['path'] for row in fuzz}=={'payload-'+str(n)+'-'+str(method)+'.zip' for n in [0,17,131089] for method in [0,8]}|{'parse_zip.rs.txt','Cargo.toml.txt'}
    post=load(BUNDLE/'fuzz-inputs/post-run.json')
    assert len(post['artifacts'])==2 and sha(BUNDLE/'fuzz-inputs/Cargo.lock.txt')==post['artifacts'][0]['sha256']
    for row in fuzz:
        retained=artifact(BUNDLE/'fuzz-inputs',row)
        if precleanup:assert sha(Path('/tmp/litchi-goal-0456/fuzz')/row['original_path'])==sha(retained)
    assert sha(BUNDLE/'fuzz-inputs/parse_zip.rs.txt')==manifests['candidate']['crates/soapberry-zip/fuzz/fuzz_targets/parse_zip.rs']
    supplements=load(ROOT/'supplement-protocol.json');assert supplements['change']==456 and supplements['status']=='frozen'
    assert set(supplements['commands'])=={'native','profiles','stack-profiles-local'}
    for name,digest in supplements['bound_files'].items():assert sha(ROOT/name)==digest
    failed_native=load(ROOT/'checks/native.json')
    assert failed_native['status']=='failed' and failed_native['exit_code']==1 and failed_native['source_unchanged']
    assert failed_native['source_before']==failed_native['source_after']==builds['candidate']['source_manifest']
    assert failed_native['argv']==supplements['commands']['native'] and failed_native['driver_sha256']==sha(ROOT/'check.py');artifact(ROOT,failed_native['log'])
    retry=load(ROOT/'supplement-r1-protocol.json');assert retry['change']==456 and retry['status']=='frozen'
    assert set(retry['commands'])=={'native-r1','profiles','stack-profiles-local'}
    for name,digest in retry['bound_files'].items():assert sha(ROOT/name)==digest
    for name,argv in retry['commands'].items():gate(ROOT,name,builds['candidate']['source_manifest'],argv)
    original_expected=(ROOT/'native-expected.json').read_bytes();correct_expected=(ROOT/'native-expected-r1.json').read_bytes()
    assert original_expected.replace(b'045db62a04565453',b'045db62a04555453')==correct_expected
    gate(ROOT,'native-r1',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0456/formal/native-r1.py'])
    native=load(ROOT/'native-proof.json');expected=load(ROOT/'native-expected-r1.json')
    assert native['status']=='pass' and native['expected_sha256']==sha(ROOT/'native-expected-r1.json')
    assert native['binary']==builds['candidate']['binaries']['external'] and native['fixture_sha256']==expected['fixture_sha256']
    assert [row['provider'] for row in native['rows']]==['bytes','range']
    for row in native['rows']:
        assert row['exit_code']==0;artifact(ROOT,row['report'])
        assert row['output']['sha256']==expected['output_sha256'] and row['output']['bytes']==expected['output_bytes']
        if precleanup:assert sha(Path(row['output']['path']))==row['output']['sha256']
    for name,script,proofname,order in [('profiles','profile.py','profile-proof.json',['baseline','candidate','candidate','baseline']),('stack-profiles-local','profile-stacks-local.py','stack-profile-local-proof.json',['baseline','candidate'])]:
        gate(ROOT,name,builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0456/formal/'+script])
        proof=load(ROOT/proofname);assert proof['status']=='pass' and len(proof['rows'])==len(order)
        for row,kind in zip(proof['rows'],order):
            assert row['build']==kind and row['binary']==builds[kind]['binaries']['normal'] and row['exit_code']==row['oracle_exit_code']==0
            for item in row['artifacts']:artifact(ROOT,item)
            if precleanup and 'raw_profile' in row:assert sha(Path(row['raw_profile']['path']))==row['raw_profile']['sha256']
    probe=load(BUNDLE/'probe-build.json');gate(BUNDLE,'probe-build',probe['source_manifest'])
    assert manifest(BUNDLE,probe['source_manifest'])==manifests['baseline']
    if precleanup:assert sha(Path(probe['binary']['path']))==probe['binary']['sha256']
    gate(BUNDLE,'pair-probe',probe['source_manifest'],['python3','-B','docs/performance/results/change-0456/probe-pairs.py'])
    pairs=load(BUNDLE/'pair-probe.json');assert pairs['binary']==probe['binary'] and pairs['driver_sha256']==sha(BUNDLE/'probe-pairs.py') and len(pairs['rows'])==6
    fixture_a='3rdparty/libreoffice-core/sd/qa/unit/data/smoketest.pptx'
    fixture_b='3rdparty/poi/test-data/slideshow/at.ecodesign.www_downloads_Vertiefungsvortrag_elektronik.pptx'
    fixtures={fixture_a:('88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44',29956),fixture_b:('4c9630fe85061c2bcb9d5638ad810fc18e8d7ae301a96d446f72b40995773578',1949517)}
    assert [(r['source']['path'],r['source_position'],r['destination']['path'],r['destination_position']) for r in pairs['rows']]==[(fixture_a,0,fixture_b,i) for i in [16,17,33]]+[(fixture_b,i,fixture_a,0) for i in [16,17,33]]
    for row in pairs['rows']:
        for side in ['source','destination']:
            fixture=row[side];assert (fixture['sha256'],fixture['bytes'])==fixtures[fixture['path']]
            if precleanup:
                actual=ROOT.parents[4]/fixture['path'];assert sha(actual)==fixture['sha256'] and actual.stat().st_size==fixture['bytes']
        assert row['source']['sha256']!=row['destination']['sha256']
        assert row['exit_code']==0 and row['stdout'].startswith('REFUSED SlideCopyPlan { kind: SharedOwner,') and not row['stderr']
        assert row['source_before']==row['source_after']==probe['source_manifest']
        assert row['argv']==[probe['binary']['path'],row['source']['path'],str(row['source_position']),row['destination']['path'],str(row['destination_position'])]
    regime=load(ROOT/'allocator-regime-protocol.json');assert regime['change']==456 and regime['status']=='frozen'
    assert regime['samples']==30 and regime['warmups']==3 and regime['cpu']==2 and regime['workers']==1
    assert regime['inherited_allocator_environment']=={} and regime['driver_sha256']==sha(ROOT/'allocator-regime.py')
    assert regime['lanes']==[dict(build=kind,mmap_threshold=threshold) for threshold in [131072,33554432] for kind in ['baseline','candidate','candidate','baseline']]
    gate(ROOT,'allocator-regime',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0456/formal/allocator-regime.py'])
    proof=load(ROOT/'allocator-regime-proof.json');assert proof['status']=='pass' and proof['protocol_sha256']==sha(ROOT/'allocator-regime-protocol.json') and proof['driver_sha256']==regime['driver_sha256']
    assert len(proof['rows'])==8 and proof['glibc']=='glibc 2.43'
    for i,(row,lane) in enumerate(zip(proof['rows'],regime['lanes'])):
        assert row['lane']==i and row['build']==lane['build'] and row['mmap_threshold']==lane['mmap_threshold']
        assert row['environment']=={'MALLOC_MMAP_THRESHOLD_':str(lane['mmap_threshold'])}
        assert row['source_before']==row['source_after']==builds['candidate']['source_manifest']
        assert row['binary']==builds[lane['build']]['binaries']['normal'] and row['exit_code']==row['oracle_exit_code']==0
        for item in row['artifacts']:artifact(ROOT,item)
        report=load(ROOT/'allocator-regime'/str(i)/'report.json');oracle.check_report(report);check_sink_consistency(report)
        assert report['samples']==30 and report['warmup']==3 and report['binary_sha256']==row['binary']['sha256']
    diagnostic=module('diagnostic_summary',ROOT/'diagnostic-summary.py');data=diagnostic.derive()
    assert data==load(ROOT/'diagnostic-summary.json') and diagnostic.render(data)==(ROOT/'diagnostic-summary.md').read_text()
    derived=module('derive',ROOT/'derive.py').derive();assert derived==load(ROOT/'measurements.json');check_markdown(ROOT)
    with tempfile.TemporaryDirectory(prefix='litchi-0456-supplements-') as directory:
        target=Path(directory)
        for name in ['protocol.json','profile-proof.json','allocation-summary.py','profile-summary.py','verify-negative.py','verify-report.py','lifecycle-oracle.py','base-verify-report.py','allocation-scopes.json']:shutil.copyfile(ROOT/name,target/name)
        for name in ['runs','profiles']:shutil.copytree(ROOT/name,target/name)
        for script,output in [('allocation-summary.py','allocation-summary.json'),('profile-summary.py','profile-summary.json'),('verify-negative.py','negative-results.json')]:
            r=subprocess.run([sys.executable,'-B',str(target/script)],capture_output=True,text=True);assert r.returncode==0,r.stderr
            assert (target/output).read_bytes()==(ROOT/output).read_bytes(),output
    if (ROOT/'final-negative-results.json').exists():
        n=load(ROOT/'final-negative-results.json');assert n['status']=='pass' and n['verifier_sha256']==sha(Path(__file__)) and len(n['results'])==2
        assert all(row['status']=='rejected' for row in n['results'])
    elif not precleanup:raise AssertionError('final sink negatives missing')
    module('transient_custody',ROOT/'verify-custody.py').check_transients(ROOT,precleanup)
    if not precleanup:
        r=load(BUNDLE/'cleanup.json');assert r['status']=='pass' and r['temporary_directory_absent'] and not Path(r['task']).exists()
        inventory=load(artifact(BUNDLE,r['inventory']));assert inventory['files']==r['files_removed'] and inventory['bytes']==r['bytes_removed']
        assert sum(row['bytes'] for row in inventory['artifacts'])==inventory['bytes'] and len(inventory['artifacts'])==inventory['files']
        pre=load(BUNDLE/'precleanup.json');assert pre['status']=='pass' and pre['exit_code']==0
        assert pre['verifier_sha256']==sha(Path(__file__)) and r['precleanup_sha256']==sha(BUNDLE/'precleanup.json')
        assert pre['driver_sha256']==sha(BUNDLE/'precleanup.py')
        assert pre['source_before']==pre['source_after']==builds['candidate']['source_manifest']
        assert pre['argv']==['/usr/bin/python3','-B',str(Path(pre['cwd'])/'docs/performance/results/change-0456/formal/verify.py'),'--precleanup']
        assert not pre['stderr'] and json.loads(pre['stdout'])['status']=='pass'
        assert r['driver_sha256']==sha(BUNDLE/'cleanup.py')
    if portable:
        proof=load(BUNDLE/'portable-verification.json');assert proof['status']=='pass' and proof['exit_code']==0 and proof['temporary_directory_absent']
        assert proof['verifier_sha256']==sha(Path(__file__)) and proof['driver_sha256']==sha(BUNDLE/'portable-verify.py')
        assert proof['argv']==['/usr/bin/python3','-B',str(Path(proof['cwd'])/'docs/performance/results/change-0456/formal/verify.py')]
        assert json.loads(proof['stdout'])['status']=='pass' and not proof['stderr'] and not Path(proof['cwd']).exists()
    return {'status':'pass','change':456,'lanes':24,'samples':720,'passed_tests':tests,'required_release_gates':len(commands),'fuzz_gates':3,'allocator_policy_diagnostic_samples':240,'review_flags':len(derived['review_flags'])}
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--precleanup',action='store_true');ap.add_argument('--portable',action='store_true');a=ap.parse_args()
    print(json.dumps(verify(a.precleanup,a.portable),sort_keys=True))
