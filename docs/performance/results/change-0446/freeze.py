#!/usr/bin/env python3
"""Freeze the validated Part-addition protocol once, before retained captures."""
import datetime,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
for tag in ('before-opc-strict','before-build','independent-fixtures','pilot-before','independent-oracle-probes'):
    assert json.loads((ROOT/'checks'/(tag+'.json')).read_text())['status']=='pass',tag
assert (ROOT.parents[3]/'crates/litchi-opc/src/content_type.rs').read_bytes()==(ROOT/'candidate/before-content_type.rs.txt').read_bytes()
assert not (ROOT/'runs').exists() and not (ROOT/'profiles').exists()
def sha(name):return hashlib.sha256((ROOT/name).read_bytes()).hexdigest()
order=[]
for phase in ('A1','B1','B2','A2'):
    repeat='R'+phase[-1]
    modes=['normal','allocator'] if repeat=='R1' else ['allocator','normal']
    shapes=['tiny','medium','large'] if repeat=='R1' else ['large','medium','tiny']
    for mode in modes:
        for shape in shapes:order.append({'phase':phase,'mode':mode,'shape':shape,'repeat':repeat})
roles={name:{'build_directory':name,'build_receipt':'checks/'+name+'-build.json','selector':'opc_part_add_plain_lifecycle','source_field':'opc_part_add','oracle_role':'plain'} for name in ('before','after')}
value={
 'schema':'litchi-0446-opc-part-add-baseline-v1','change':446,'status':'frozen',
 'frozen_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
 'purpose':'before/after owned content-type Part-name handoff on the plain-source lifecycle',
 'cpu':2,'workers':1,'samples':30,'warmups':3,'repeats':2,'modes':['normal','allocator'],
 'shapes':{'tiny':64,'medium':1024,'large':4096},
 'roles':roles,
 'order':order,'matrix':{'reports':24,'retained_samples':720,'reports_per_phase':6,'samples_per_report':30},
 'timing_scope':{'source_modes':'both builds use plain OwnedSource and the unchanged timed harness body','inside':['source-backed catalog opening','one Part and root relationship plan construction','consuming sequential publication including catalog drop'], 'outside':['source wrapper and prepared input/payload','sink setup','digest finalization, source snapshots, full artifact oracles and report assembly'], 'data_path_sha256':sha('data-path.md')},
 'profiles':{'roles':['before','after'],'symbolization':'perf report --no-inline -g none; perf script --no-inline','shape':'large','mode':'normal','kinds':['stat','record'],'events':['cycles:u','instructions:u','branches:u','branch-misses:u','L1-dcache-load-misses:u'],'record_frequency_hz':999,'call_graph':'fp,127','scope':'whole fresh process including fixtures, untimed gates, warmups, timed samples and report hashing; Python oracle afterward'},
 'oracle':{'verifier_path':'oracle/verify-report.py','verifier_sha256':sha('oracle/verify-report.py'),'fixtures_verifier_sha256':sha('oracle/fixtures.py'),'protocol_path':'oracle/protocol.json','protocol_sha256':sha('oracle/protocol.json'),'role_map':{'before':'plain','after':'plain'},'argv':['{python}','-B','{verifier}','--report','{report}','--mode','{mode}','--shape','{shape}','--role','{role}'],'contract':'actual ZIP fixture identities, independent payload/XML/raw-record verification, report/catalog hashes, source/sink alignment and allocation balance; producer refusal gates are separate Rust checks'},
 'analysis':{'p50':'arithmetic midpoint of central observations','p95_p99':'nearest rank','bootstrap_95':'2000 deterministic resamples per report','bootstrap_seed':4460301,'repeat_review_trigger_percent':5,'scope':'descriptive baseline; allocator elapsed time is not a latency claim'},
 'claims':{'allowed':['before/after allocation observations; normal latency only if its separate practical gate passes'], 'withheld':['universal Office speedup','native Office owner creation','cold/range or scaling','bounded total memory','causal attribution from whole-process profiles']},
 'acceptance':{'allocation_calls_reduction_percent':5,'shapes':['medium','large'],'both_repeats_required':True,'normal_p50_claim_reduction_percent':3,'regression_review_trigger_percent':5,'correctness_required':True,'candidate_before_sha256':sha('candidate/before-content_type.rs.txt'),'candidate_after_sha256':sha('candidate/after-content_type.rs.txt')},
 'freeze_scope':'baseline fixture/pilot validation preceded this freeze; production edit and all retained captures/profiles must follow it'
}
with (ROOT/'protocol.json').open('x') as stream:stream.write(json.dumps(value,indent=2)+'\n')
print(json.dumps({'status':'frozen','frozen_utc':value['frozen_utc'],'protocol_sha256':sha('protocol.json')}))
