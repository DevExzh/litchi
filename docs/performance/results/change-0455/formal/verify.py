#!/usr/bin/env python3
"""Replay the exact matched capture, release gates, derivation and cleanup custody."""
import argparse,contextlib,hashlib,importlib.util,io,json,shutil,subprocess,sys,tempfile
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
    assert r['source_before']==r['source_after'];manifest(base,r['source_after']);artifact(base,r['log'])
    assert sha(base/'check.py')==r['driver_sha256']
    if expected_source:assert r['source_after']==expected_source,name
    if argv:assert r['argv']==argv,name
    if 'test' in r['argv']:assert r['passed_tests']>0 and r['failed_tests']==0
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
    with tempfile.TemporaryDirectory(prefix='litchi-0455-render-') as directory:
        target=Path(directory);shutil.copyfile(base/'protocol.json',target/'protocol.json');shutil.copytree(base/'runs',target/'runs')
        renderer=module('markdown_renderer',base/'derive.py');renderer.ROOT=target
        with contextlib.redirect_stdout(io.StringIO()):renderer.main()
        for name in ['measurements.json','measurements.md']:assert (base/name).read_bytes()==(target/name).read_bytes(),name+' differs from regenerated output'
def check_supplements():
    with tempfile.TemporaryDirectory(prefix='litchi-0455-supplements-') as directory:
        target=Path(directory)
        for name in ['protocol.json','profile-proof.json','allocation-summary.py','profile-summary.py','verify-negative.py','verify-report.py','lifecycle-oracle.py','base-verify-report.py','allocation-scopes.json']:
            shutil.copyfile(ROOT/name,target/name)
        for name in ['runs','profiles']:shutil.copytree(ROOT/name,target/name)
        for script,output in [('allocation-summary.py','allocation-summary.json'),('profile-summary.py','profile-summary.json'),('verify-negative.py','negative-results.json')]:
            r=subprocess.run([sys.executable,'-B',str(target/script)],capture_output=True,text=True);assert r.returncode==0,r.stderr
            assert (target/output).read_bytes()==(ROOT/output).read_bytes(),output
def verify(precleanup=False,portable=False):
    assert not (precleanup and portable)
    assert sha(ROOT/'final-gate-bindings.json')=='3e890594f9f866f35afd317ad4935d9a255210f2f6974778bc849f5bcd121883'
    bindings=load(ROOT/'final-gate-bindings.json')
    for name,digest in bindings['files'].items():assert sha(ROOT/name)==digest,name
    assert sha(BUNDLE/'supplement-bindings.json')=='1ddb90268ab985c1ee2d12a9c6eaeb8bf239464f6a7ead2e2102a7f390a347da'
    for name,digest in load(BUNDLE/'supplement-bindings.json')['files'].items():assert sha(BUNDLE/name)==digest,name
    protocol=load(ROOT/'protocol.json');assert protocol['samples']==30 and protocol['warmups']==3 and len(protocol['lanes'])==24
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
    changed=[p for p,h in manifests['baseline'].items() if manifests['candidate'].get(p)!=h]
    assert set(manifests['baseline'])==set(manifests['candidate'])
    assert set(changed)=={'crates/soapberry-zip/src/preserve.rs','crates/soapberry-zip/tests/preservation_zip64_promotion.rs'},changed
    before=(ROOT/'baseline-preserve.rs.txt').read_bytes();after=(ROOT/'candidate-preserve.rs.txt').read_bytes()
    assert before.replace(b'const COPY_CHUNK_SIZE: usize = 32 * 1024;',b'const COPY_CHUNK_SIZE: usize = 64 * 1024;',1)==after,'production delta is not exactly buffer size'
    for kind,raw in [('baseline',before),('candidate',after)]:assert hashlib.sha256(raw).hexdigest()==manifests[kind]['crates/soapberry-zip/src/preserve.rs']
    test_before=(ROOT/'baseline-zip64-test.rs.txt').read_bytes();test_after=(ROOT/'candidate-zip64-test.rs.txt').read_bytes()
    assert test_before.replace(b'const COPY_CHUNK_SIZE: usize = 32 * 1024;',b'const COPY_CHUNK_SIZE: usize = 64 * 1024;',1)==test_after
    for kind,raw in [('baseline',test_before),('candidate',test_after)]:assert hashlib.sha256(raw).hexdigest()==manifests[kind]['crates/soapberry-zip/tests/preservation_zip64_promotion.rs']
    for kind,b in builds.items():
        gate(ROOT,kind+'-build',b['source_manifest'])
        for binary in b['binaries'].values():
            p=Path(binary['path'])
            if precleanup:assert p.is_file()
            if p.exists():assert p.stat().st_size==binary['bytes'] and sha(p)==binary['sha256']
    preliminary_protocol=load(BUNDLE/'protocol.json')
    for name,digest in preliminary_protocol['bound_files'].items():assert sha(BUNDLE/name)==digest,name
    preliminary_build=load(BUNDLE/'baseline-build.json');gate(BUNDLE,'baseline-build',preliminary_build['source_manifest'])
    if precleanup:
        for binary in preliminary_build['binaries'].values():assert sha(Path(binary['path']))==binary['sha256'] and Path(binary['path']).stat().st_size==binary['bytes']
    preliminary_oracle=module('preliminary_oracle',BUNDLE/'verify-report.py')
    assert {p.name for p in (BUNDLE/'runs').iterdir()}=={'0','1','2','3','16','17'}
    for directory in (BUNDLE/'runs').iterdir():
        r=load(directory/'receipt.json');assert r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0
        assert r['source_before']==r['source_after']==preliminary_build['source_manifest'] and r['source_unchanged']
        assert r['protocol_sha256']==sha(BUNDLE/'protocol.json') and r['capture_sha256']==sha(BUNDLE/'capture.py')
        for item in r['artifacts'].values():artifact(BUNDLE,item)
        assert r['lane_definition']==preliminary_protocol['lanes'][int(directory.name)]
        capture_identity(BUNDLE,'docs/performance/results/change-0455',r,preliminary_build,int(directory.name))
        preliminary_oracle.check_report(load(directory/'report.json'))
    oracle=module('oracle',ROOT/'verify-report.py')
    assert {p.name for p in (ROOT/'runs').iterdir()}=={str(i) for i in range(24)}
    for i,lane in enumerate(protocol['lanes']):
        d=ROOT/'runs'/str(i);r=load(d/'receipt.json');report=load(d/'report.json');build=builds[lane['build']]
        assert r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0
        assert r['lane']==i and r['lane_definition']==lane and r['source_before']==r['source_after'] and r['source_unchanged']
        assert r['source_before']==builds['baseline' if i in [0,1,2,3,16,17] else 'candidate']['source_manifest']
        manifest(ROOT,r['source_before']);assert r['protocol_sha256']==sha(ROOT/'protocol.json') and r['capture_sha256']==sha(ROOT/'capture.py')
        artifact(ROOT,r['build_manifest']);assert r['binary']==build['binaries'][lane['instrumentation']]
        for item in r['artifacts'].values():artifact(ROOT,item)
        assert set(r['artifacts'])=={'report.json','resource.log','workload.log','oracle.log'}
        assert report['binary_sha256']==r['binary']['sha256'] and report['binary_bytes']==r['binary']['bytes']
        assert report['source_revision']==build['revision'] and report['corpus']==lane['corpus'] and report['provider']==lane['provider']
        original_directory=Path(r['cwd'])/'docs/performance/results/change-0455/formal/runs'/str(i)
        expected=['taskset','-c','2','/usr/bin/time','-v','-o',str(original_directory/'resource.log'),r['binary']['path'],'provider-lifecycle','--corpus',lane['corpus'],'--provider',lane['provider'],'--samples','30','--warmup','3','--source-revision',build['revision'],'--output',str(original_directory/'report.json')]
        if lane['provider']=='range':expected+=['--max-range','65536','--delay-us','200','--transfer-bytes-per-second','26214400','--transfer-delay-policy','separate-sleeps']
        assert r['argv']==expected,'capture command differs'
        assert report['samples']==30 and report['warmup']==3 and report['checked_iteration_count']==33
        expected_instrumentation='none' if lane['instrumentation']=='normal' else 'system_allocator_operation_scoped'
        assert report['instrumentation']==expected_instrumentation
        for key,value in protocol['corpora'][lane['corpus']].items():assert report[key]==value
        capture_identity(ROOT,'docs/performance/results/change-0455/formal',r,build,i)
        oracle.check_report(report);check_sink_consistency(report)
    commands=module('commands',ROOT/'run-checks.py').COMMANDS
    tests=0
    for name,argv in commands.items():r=gate(ROOT,name,builds['candidate']['source_manifest'],argv);tests+=r.get('passed_tests',0)
    for name,argv in module('fuzz_commands',ROOT/'run-fuzz.py').COMMANDS.items():gate(ROOT,name,builds['candidate']['source_manifest'],argv)
    fuzz_inputs=load(BUNDLE/'fuzz-inputs/manifest.json')
    for row in fuzz_inputs:
        relative=Path(row['path']);retained=BUNDLE/'fuzz-inputs'/(relative.name+'.txt' if relative.suffix in ['.rs','.toml'] else relative.name)
        assert sha(retained)==row['sha256'] and retained.stat().st_size==row['bytes']
        if precleanup:assert sha(Path('/tmp/litchi-goal-0455/fuzz')/relative)==row['sha256']
    assert sha(BUNDLE/'fuzz-inputs/parse_opc.rs.txt')==manifests['candidate']['crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs']
    gate(ROOT,'native',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0455/formal/native.py'])
    native=load(ROOT/'native-proof.json');expected=load(ROOT/'native-expected.json')
    assert native['status']=='pass' and native['expected_sha256']==sha(ROOT/'native-expected.json')
    assert native['binary']==builds['candidate']['binaries']['external'] and native['fixture_sha256']==expected['fixture_sha256']
    assert [row['provider'] for row in native['rows']]==['bytes','range']
    for row in native['rows']:
        assert row['exit_code']==0;artifact(ROOT,row['report'])
        assert row['output']['sha256']==expected['output_sha256'] and row['output']['bytes']==expected['output_bytes']
        if precleanup:assert sha(Path(row['output']['path']))==row['output']['sha256']
    assert sha(ROOT/'profile-stacks.py')=='673790beb512e66d837823b685fabcc05576116cd53c57dad9480edf319d8469'
    gate(ROOT,'profiles',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0455/formal/profile.py'])
    profiles=load(ROOT/'profile-proof.json');assert profiles['status']=='pass' and profiles['order']==['baseline','candidate','candidate','baseline']
    assert len(profiles['rows'])==4
    for row,kind in zip(profiles['rows'],profiles['order']):
        assert row['build']==kind and row['binary']==builds[kind]['binaries']['normal'] and row['exit_code']==row['oracle_exit_code']==0
        for item in row['artifacts']:artifact(ROOT,item)
    assert sha(ROOT/'checks/stack-profiles.json')=='5ce93bed4fe991ae9284bff5e8c3709b261d2d6dc2807c26f794535b27caa2ea'
    failed=load(ROOT/'checks/stack-profiles.json');assert failed['status']=='failed' and failed['exit_code']==1 and failed['source_unchanged']
    assert failed['argv']==['python3','-B','docs/performance/results/change-0455/formal/profile-stacks.py'] and failed['driver_sha256']==sha(ROOT/'check.py')
    assert failed['source_before']==failed['source_after']==builds['candidate']['source_manifest'];artifact(ROOT,failed['log'])
    assert sha(ROOT/'profile-stacks-local.py')=='449f33b188309b2abac4f287010985a18ffd6b293832017d1173f7741fbf3df4'
    gate(ROOT,'stack-profiles-local',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0455/formal/profile-stacks-local.py'])
    stacks=load(ROOT/'stack-profile-local-proof.json');assert stacks['status']=='pass' and len(stacks['rows'])==2
    for row,kind in zip(stacks['rows'],['baseline','candidate']):
        assert row['build']==kind and row['binary']==builds[kind]['binaries']['normal'] and row['exit_code']==row['oracle_exit_code']==0
        for item in row['artifacts']:artifact(ROOT,item)
        if precleanup:assert sha(Path(row['raw_profile']['path']))==row['raw_profile']['sha256']
    assert sha(ROOT/'allocator-regime-protocol.json')=='ad97c873af86502fd162f19d147a818443ee785ed144003e1422e18e8511f42f'
    regime=load(ROOT/'allocator-regime-protocol.json');assert sha(ROOT/'allocator-regime.py')==regime['driver_sha256']
    gate(ROOT,'allocator-regime',builds['candidate']['source_manifest'],['python3','-B','docs/performance/results/change-0455/formal/allocator-regime.py'])
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
    derived=module('derive',ROOT/'derive.py').derive();assert derived==load(ROOT/'measurements.json'),'derived results differ'
    confirmation=ROOT/'confirmation';cp=load(confirmation/'protocol.json')
    assert len(cp['lanes'])==8 and cp['samples']==30 and cp['warmups']==3
    assert [lane['build'] for lane in cp['lanes']]==['baseline','candidate','candidate','baseline','baseline','candidate','candidate','baseline']
    for name,digest in cp['bound_files'].items():assert sha(confirmation/name)==digest,name
    assert {p.name for p in (confirmation/'runs').iterdir()}=={str(i) for i in range(8)}
    for i,lane in enumerate(cp['lanes']):
        assert lane['provider']=='bytes' and lane['corpus']=='media-rich' and lane['instrumentation']=='normal'
        directory=confirmation/'runs'/str(i);r=load(directory/'receipt.json');report=load(directory/'report.json')
        assert r['status']=='pass' and r['exit_code']==r['oracle_exit_code']==0 and r['lane_definition']==lane
        assert r['source_before']==r['source_after']==builds['candidate']['source_manifest'] and r['source_unchanged']
        manifest(confirmation,r['source_before'])
        assert r['binary']==builds[lane['build']]['binaries']['normal']
        assert r['protocol_sha256']==sha(confirmation/'protocol.json') and r['capture_sha256']==sha(confirmation/'capture.py')
        for item in r['artifacts'].values():artifact(confirmation,item)
        assert report['samples']==30 and report['warmup']==3 and report['binary_sha256']==r['binary']['sha256']
        capture_identity(confirmation,'docs/performance/results/change-0455/formal/confirmation',r,builds[lane['build']],i)
        oracle.check_report(report);check_sink_consistency(report)
    assert module('confirmation_derive',confirmation/'derive.py').derive()==load(confirmation/'measurements.json')
    check_markdown(ROOT);check_markdown(confirmation)
    assert sha(ROOT/'diagnostic-summary.py')=='b819cf378b32829a7652a412b0019b13f36c01a7c8ce0d16d1b274466319224c'
    diagnostic=module('diagnostic_summary',ROOT/'diagnostic-summary.py');data=diagnostic.derive()
    assert data==load(ROOT/'diagnostic-summary.json') and diagnostic.render(data)==(ROOT/'diagnostic-summary.md').read_text()
    check_supplements()
    if (ROOT/'final-negative-results.json').exists():
        negatives=load(ROOT/'final-negative-results.json');assert negatives['status']=='pass' and negatives['verifier_sha256']==sha(Path(__file__))
        assert len(negatives['results'])==2 and all(row['status']=='rejected' for row in negatives['results'])
    elif not precleanup:raise AssertionError('final sink negatives missing')
    if not precleanup:
        r=load(BUNDLE/'cleanup.json');assert r['status']=='pass' and r['temporary_directory_absent'] and not Path(r['task']).exists()
        inventory=load(artifact(BUNDLE,r['inventory']));assert inventory['files']==r['files_removed'] and inventory['bytes']==r['bytes_removed']
        assert sum(row['bytes'] for row in inventory['artifacts'])==inventory['bytes'] and len(inventory['artifacts'])==inventory['files']
        pre=load(BUNDLE/'precleanup.json');assert pre['status']=='pass' and pre['exit_code']==0
        assert pre['verifier_sha256']==sha(Path(__file__)) and r['precleanup_sha256']==sha(BUNDLE/'precleanup.json')
        assert r['driver_sha256']==sha(BUNDLE/'cleanup.py')
    if portable:
        proof=load(BUNDLE/'portable-verification.json');assert proof['status']=='pass' and proof['exit_code']==0 and proof['temporary_directory_absent']
        assert proof['verifier_sha256']==sha(Path(__file__)) and proof['driver_sha256']==sha(BUNDLE/'portable-verify.py')
        assert proof['argv']==['/usr/bin/python3','-B',str(Path(proof['cwd'])/'docs/performance/results/change-0455/formal/verify.py')]
        assert json.loads(proof['stdout'])['status']=='pass' and not proof['stderr'] and not Path(proof['cwd']).exists()
    return {'status':'pass','change':455,'lanes':24,'samples':720,'passed_tests':tests,'required_release_gates':len(commands),'fuzz_gates':3,'review_flags':len(derived['review_flags']),'confirmation_samples':240,'allocator_policy_diagnostic_samples':240}
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--precleanup',action='store_true');ap.add_argument('--portable',action='store_true');a=ap.parse_args()
    result=verify(a.precleanup,a.portable);print(json.dumps(result,sort_keys=True))
