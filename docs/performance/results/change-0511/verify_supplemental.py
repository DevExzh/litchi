"""Validate diagnostic comparisons and source-bound correctness gates."""
import hashlib,json,re
from tools.summarize_crud_baseline import _validate_elapsed
from tools.validate_perf_corpus_binding import validate_binding

def load(path):return json.loads(path.read_text())
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def ir(path):return int(re.search(r'^summary: (\d+)',path.read_text(),re.M)[1])
def validate(here):
    profiles={}
    for stage in ['before','after']:
        directory=here/stage
        build=load(here/'before/build-receipt.json' if stage=='before' else here/'build-receipt.json')
        receipt=load(directory/'cfb-profile-receipt.json')
        assert receipt['exit_code']==0 and receipt['source_unchanged']
        assert receipt['binary_sha256']==build['binary_sha256']
        assert receipt['source_manifest_sha256']==build['source_manifest_sha256']
        assert '--toggle-collect=*run_cfb_open' in receipt['command']
        for name,digest in receipt['artifacts'].items():assert sha(directory/name)==digest
        report=load(directory/'cfb-profile-report.json')
        validate_binding(report,load(directory/'cfb-profile-catalog.json'))
        assert report['binary_identity']['binary_sha256']==build['binary_sha256']
        assert report['configuration']['samples_per_case']==5 and report['configuration']['warmup_iterations_per_case']==0
        assert len(report['results'])==1
        row=report['results'][0];_validate_elapsed(row,5,stage+'cfb-profile')
        reference=next(r for r in load(directory/'guard-r1-report.json')['results'] if r['corpus']['shape']=='few-large')
        for field in ['case','corpus','sink']:assert row[field]==reference[field]
        assert row.get('output_sha256')==reference.get('output_sha256')
        tree=(directory/'cfb-profile-inclusive.txt').read_text()
        assert 'OleFile<R>::open (5x)' in tree
        total=ir(directory/'cfb-profile.out');assert total>0
        constructor=int(re.search(r'([\d,]+) \([^\n]*\)  >   \?\?\?:litchi_cfb::file::OleFile<R>::open \(5x\)',tree)[1].replace(',',''))
        profiles[stage]={'runner_ir':total,'constructor_inclusive_ir':constructor}
    original=here/'initial-push'
    rejection=load(original/'rejection.json')
    assert rejection['before_ir']==ir(here/'before/callgrind.out')
    assert rejection['after_ir']==ir(original/'after/callgrind.out')
    assert rejection['after_ir']>rejection['before_ir']
    initial_build=load(original/'build-receipt.json')
    assert initial_build['exit_code']==0 and initial_build['source_unchanged']
    assert sha(original/'source-manifest.json')==initial_build['source_manifest_sha256']
    assert sha(original/'build.log')==initial_build['log_sha256']
    receipt=load(original/'after/callgrind-receipt.json')
    assert receipt['binary_sha256']==initial_build['binary_sha256']
    for name,digest in receipt['artifacts'].items():assert sha(original/'after'/name)==digest
    gates=load(here/'gates.json')
    final_manifest=sha(here/'source-manifest.json');total_passed=0;total_ignored=0
    assert set(gates['source_bound_receipts'])=={'fmt','ole2-tests','ole2-doctests','cfb-no-default-tests','cfb-clippy','cfb-rustdoc'}
    for name,receipt in gates['source_bound_receipts'].items():
        assert receipt==load(here/f'{name}-receipt.json')
        assert receipt['exit_code']==0 and receipt['source_unchanged']
        assert receipt['source_manifest_sha256']==final_manifest
        assert receipt['log_sha256']==sha(here/f'{name}.log')
        if name in ['ole2-tests','ole2-doctests','cfb-no-default-tests']:
            counts=re.findall(r'test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored;', (here/f'{name}.log').read_text())
            assert counts and all(int(r[1])==0 for r in counts)
            passed=sum(int(r[0]) for r in counts);ignored=sum(int(r[2]) for r in counts)
            assert passed==gates['test_counts'][name]['passed']
            assert ignored==gates['test_counts'][name]['ignored']
            total_passed+=passed;total_ignored+=ignored
    assert total_passed==gates['total_passed'] and total_ignored==gates['total_ignored']
    assert total_passed>4000
    for name,receipt in gates['additional_checks'].items():
        assert receipt['exit_code']==0 and receipt['log_sha256']==sha(here/name)
    return {'cfb_guard_runner_ir':profiles,'cfb_guard_runner_ir_change_percent':(profiles['after']['runner_ir']/profiles['before']['runner_ir']-1)*100,'initial_push_candidate_rejected':True,'test_executions_passed':total_passed,'test_executions_ignored':total_ignored,'scope':'CFB profile includes five opens, timers, file-size oracles, drops and result construction; fixture generation excluded; all instrumented timings excluded'}
