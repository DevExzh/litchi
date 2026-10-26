"""Read-only root verification of the frozen C2a gate; no execution/restart."""
import argparse
import datetime
import hashlib
import importlib.util
import json
import pathlib
import subprocess

BASE = pathlib.Path('/home/zhuhe/code/litchi-worktrees/.coord')
PACKET = BASE / 'final-gate-C2a-33a0a63d35-20261026'
STATE = BASE / 'final-gate-C2a-33a0a63d35-state-20261026'
DRIVER = BASE / 'final-gate-C2a-33a0a63d35-driver-20261026/final_gate_v6.py'
DRIVER_HASH = 'd9ef88dcfb048bae60ac2a4f1cca2072f9b99ff3d7383f25d340a7df7469aeee'
PLAN_HASH = 'e76fed802a2b87072285369dfae1286e90bde5107d919439fa2d74853bf24ec2'
UNIT = 'litchi-unified-final65-c2a-33a0-v1.service'

def require(ok, message):
    if not ok:
        raise ValueError(message)

def identity(path):
    path = pathlib.Path(path)
    digest = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024*1024), b''):
            digest.update(block)
    return {'path': str(path), 'bytes': path.stat().st_size, 'sha256': digest.hexdigest()}

def check_file(path, size, digest):
    item = identity(path)
    require(item['bytes'] == size and item['sha256'] == digest, f'file identity: {path}')
    return item

def check_shape(rows, plan, terminal):
    phases = ['compiler', 'full53', 'runtime']
    require(set(rows) == set(phases), 'phase set mismatch')
    incomplete = False
    for phase in phases:
        expected = plan['phases'][phase]['steps']
        actual = rows[phase]
        require(isinstance(actual, list) and len(actual) <= len(expected), f'{phase} count overflow')
        require(not incomplete or not actual, 'later phase started before prior phase complete')
        if terminal:
            require(len(actual) == len(expected), f'{phase} incomplete')
        incomplete |= len(actual) != len(expected)
        for i, row in enumerate(actual):
            step = expected[i]
            require(row['index'] == i and row['total'] == len(expected) and row['name'] == step['name'], f'{phase}:{i} sequence mismatch')
            require(row['status'] == 'passed' and row['exit'] == 0 and row['effective_exit'] == 0, f'{phase}:{i} did not pass')
            argv = step['argv'] if step['kind'] == 'argv' else ['bash','-c',step['command']]
            require(row['command'] == argv, f'{phase}:{i} command mismatch')

def verify_row(row, phase, plan):
    i = row['index']; step = plan['phases'][phase]['steps'][i]
    label = f'{phase}:{i}:{row["name"]}'
    expected_source = plan['runner_source']; locks = expected_source['cargo_locks']
    require(row['source_after'] == expected_source, f'{label} source changed')
    require(row['cargo_locks_before'] == locks == row['cargo_locks_after'], f'{label} lock changed')
    require(not row.get('resource_failure') and not row.get('provenance_failure'), f'{label} failure field')
    require(not row['target_live_guard_errors'], f'{label} guard errors')
    env = row['environment']
    required_env = {'CARGO_TARGET_DIR':plan['target'], 'CARGO_NET_OFFLINE':'true','LITCHI_MEM_MAX':'20G','CARGO_BUILD_JOBS':'1','RUST_TEST_THREADS':'1','CARGO_INCREMENTAL':'0','LITCHI_GATE_ATTEMPT':f'final-v5-{phase}-{i:02d}'}
    if 'RUSTDOCFLAGS' in step['env']:
        required_env['RUSTDOCFLAGS'] = step['env']['RUSTDOCFLAGS']
    require(env == required_env, f'{label} recorded environment mismatch')
    pre = row['preflight']; live = row['target_live_guard']; post = row['post_exit']
    require(pre['target_bytes'] <= 18_000_000_000 and pre['free_bytes'] >= 12_000_000_000 and not pre['builders'], f'{label} preflight resource/process guard')
    require(live['samples'] > 0 and live['live_target'] == 19_000_000_000 and live['max_sampled_bytes'] < 19_000_000_000, f'{label} live target guard')
    require(row['target_bytes_after'] <= 20_000_000_000 and row['free_bytes_after'] >= 12_000_000_000, f'{label} post resources')
    require(post['target_bytes'] == row['target_bytes_after'] and len(post['samples']) >= 3, f'{label} post-exit sampling')
    require(all(not x['owned'] and not x['builders'] and x['target_bytes'] <= 20_000_000_000 for x in post['samples']), f'{label} post-exit holders/resources')
    log = check_file(row['log'], row['log_bytes'], row['log_sha256'])
    trace = check_file(row['cargo_target_trace_file'], row['cargo_target_trace_bytes'], row['cargo_target_trace_sha256'])
    parsed = []
    for line in pathlib.Path(trace['path']).read_text().splitlines():
        fields = line.split('\t')
        require(len(fields) == 3 and fields[0].isdigit(), f'{label} malformed target trace')
        parsed.append(dict(zip(['pid','cargo_command','target'], fields)))
    require(parsed == row['cargo_target_trace'], f'{label} trace projection mismatch')
    targets = {x['target'] for x in parsed}
    require(targets == set(row['cargo_target_sizes']), f'{label} missing target sizes')
    for target, size in row['cargo_target_sizes'].items():
        require(pathlib.Path(target).resolve().is_relative_to(pathlib.Path(plan['target']).resolve()), f'{label} target escaped')
        require(isinstance(size['bytes'],int) and size['bytes'] <= 20_000_000_000 and not size.get('error'), f'{label} nested target invalid')
    return {'phase':phase,'index':i,'name':row['name'],'log':log,'trace':trace,'target_bytes_after':row['target_bytes_after']}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--allow-partial', action='store_true')
    parser.add_argument('--output', type=pathlib.Path, required=True)
    args = parser.parse_args()
    require(not args.output.exists(), 'output must be fresh')
    require(identity(DRIVER)['sha256'] == DRIVER_HASH, 'driver digest changed')
    require(identity(PACKET/'plan.json')['sha256'] == PLAN_HASH, 'plan digest changed')
    plan = json.loads((PACKET/'plan.json').read_text())
    spec = importlib.util.spec_from_file_location('reviewed_frozen_gate_driver', DRIVER)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
    # verify_plan(go=False) validates source/tool/command derivation without GO,
    # empty-state checks, Cargo, native children or archive generation.
    module.verify_plan(pathlib.Path(plan['source']['root']), PACKET, plan, go=False, expected_plan_sha256=PLAN_HASH)
    rows_bytes = (STATE/'rows.json').read_bytes(); rows = json.loads(rows_bytes)
    terminal = not args.allow_partial
    check_shape(rows, plan, terminal)
    verified = [verify_row(row,phase,plan) for phase in ['compiler','full53','runtime'] for row in rows[phase]]
    service_raw = subprocess.check_output(['systemctl','--user','show',UNIT,'--property=ActiveState,SubState,MainPID,ExecMainStatus,Result'],text=True)
    service = dict(line.split('=',1) for line in service_raw.splitlines())
    result_identity = None
    if terminal:
        require(service['ActiveState'] == 'inactive' and service['MainPID'] == '0' and service['ExecMainStatus'] == '0' and service['Result'] == 'success', 'gate service not successfully terminal')
        result = json.loads((STATE/'result.json').read_text())
        require(result['status'] == 'passed' and result['rows'] == rows, 'terminal result differs from verified rows')
        require(result.get('officecli_preservation') and rows['full53'][37].get('officecli_preservation') == result['officecli_preservation'], 'missing corpus preservation barrier')
        result_identity = identity(STATE/'result.json')
    else:
        require(service['ActiveState'] == 'active' and int(service['MainPID']) > 0, 'partial wait requires confirmed live service')
    # Snapshot only the bytes actually checked, not a later concurrent rows read.
    output = {'schema':'litchi.root-C2a-gate-row-verification.v1','utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'status':'TERMINAL_ROWS_VERIFIED_PENDING_NESTED_AND_CORPUS_POSTCHECK' if terminal else 'PARTIAL_ROWS_VERIFIED_LIVE_GATE','source':plan['runner_source'],'plan_sha256':PLAN_HASH,'verifier':identity(pathlib.Path(__file__)),'rows_snapshot_sha256':hashlib.sha256(rows_bytes).hexdigest(),'rows_count':len(verified),'rows':verified,'service':service,'result':result_identity,'limitations':['Does not accept native qualification or raw API census.','Nested chunk completeness/logs and preserved OfficeCLI outcome bytes require their separate root checks.']}
    args.output.mkdir(parents=True,exist_ok=False)
    (args.output/'rows-snapshot.json').write_bytes(rows_bytes)
    (args.output/'verification.json').write_text(json.dumps(output,indent=2)+'\n')
    print(json.dumps({'status':output['status'],'rows':len(verified),'receipt':identity(args.output/'verification.json')},indent=2))

if __name__ == '__main__':
    main()
