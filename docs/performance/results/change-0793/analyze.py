"""Fail-closed 0793 evidence replay; failed experimental gates remain data."""
import argparse
from collections import Counter
import json
from pathlib import Path
import re
import subprocess
import custody as c

P, ROOT = c.P, c.ROOT
SHAPES = ['tiny', 'medium', 'large', 'vendor', 'unicode-vendor']
VARIANTS = ['control', 'profile', 'allocation', 'profile-allocation']
OWNER = re.compile(r'^namespace_uri_probe::capture_region_0793::h[0-9a-f]+(?: \(.*\))?$')


def artifact(a):
    raw = Path(a['path'])
    origin = c.read(P/'origin.json')
    if raw.is_absolute():
        path = ROOT / raw.relative_to(origin['main'])
    else:
        path = P / raw
    assert path.stat().st_size == a['bytes'] and c.sha(path) == a['sha256'], path
    return path


def receipts(lane, count, label):
    done = c.read(P/lane/'complete.json')
    assert done[label] == count
    assert artifact(done['receipts']) == P/lane/'receipts.json'
    if 'source' in done:
        assert artifact(done['source']) == P/'build/source.json'
    rows = c.read(P/lane/'receipts.json')
    assert len(rows) == count
    for row in rows:
        assert row['exit_code'] == 0 and row['ended'] >= row['started']
        artifact(row['log'])
    assert all(a['ended'] <= b['started'] for a,b in zip(rows,rows[1:]))
    return rows


def parse_stacks(path):
    rows = Counter()
    for line in path.read_text().splitlines():
        stack, n = line.rsplit(' ', 1)
        assert n.isdecimal() and int(n) > 0
        rows[stack] += int(n)
    return rows


def counter(a):
    assert a['status'] == 'measured' and a['scope'] == 'operation_global_system_allocator'
    assert a['failed_allocation_calls'] == 0
    return {**{k:a[k] for k in ['allocation_calls','reallocation_calls','deallocation_calls','allocated_bytes','deallocated_bytes']},
            'net_live_bytes':a['live_bytes_after']-a['live_bytes_before'],
            'peak_above_entry_bytes':a['region_peak_live_bytes']-a['live_bytes_before']}


def analyze():
    origin = c.read(P/'origin.json')
    plan = c.read(P/'plan.json')
    assert plan['shapes'] == SHAPES and plan['cpu'] == 12
    assert plan['controls']['orders'] == [VARIANTS, list(reversed(VARIANTS))]
    assert plan['heaptrack']['orders'] == [SHAPES,list(reversed(SHAPES))]
    assert 'allocation_calls plus reallocation_calls' in plan['qualification'][4]
    for name, digest in c.read(P/'build/frozen-inputs.json').items():
        assert c.sha(P/name) == digest, name
    inputs = c.read(P/'architecture-inputs.json')
    assert len(inputs) == 35
    for name,digest in inputs.items():
        assert c.sha(ROOT/name) == digest, name
    inherit = c.read(P/'inheritance.json')
    prior = P/inherit['prior_packet']
    assert c.sha(prior/'seal.json') == inherit['prior_seal']
    seal = c.read(prior/'seal.json')
    def inherited(name):
        path = prior/name
        assert c.sha(path) == seal['files'][name], name
        return c.read(path)
    source = c.read(P/'build/source.json')
    assert artifact(inherit['production_source']).resolve() == (prior/'build-after/source.json').resolve()
    assert source['files'] == inherited('build-after/source.json')['files']
    assert len(source['files']) == 9196 and source['revision'] == origin['base']
    assert c.source()['files'] == source['files']
    subprocess.run(['git','merge-base','--is-ancestor',origin['base'],'HEAD'],cwd=ROOT,check=True)
    base_seal = subprocess.check_output(['git','show',origin['base']+':docs/performance/results/change-0792/seal.json'],cwd=ROOT)
    import hashlib
    assert hashlib.sha256(base_seal).hexdigest() == inherit['prior_seal']
    assert c.sha(P/'workspace-Cargo.lock') == c.sha(prior/'workspace-Cargo.lock') == seal['files']['workspace-Cargo.lock']
    for name,digest in inherit['probe_reference_files'].items():
        assert c.sha(prior/'probe-src'/name) == digest == seal['files']['probe-src/'+name]
        if name not in ['Cargo.toml.template','src/main.rs']:
            assert c.sha(P/'probe-src'/name) == digest
    build = c.read(P/'build/receipt.json')
    assert build['rows'] == c.read(P/'build/commands.json')
    assert build['probe'] == c.read(P/'build/probe.json')
    for name,digest in build['probe'].items(): assert c.sha(P/name) == digest
    assert artifact(build['source']) == P/'build/source.json'
    assert list(build['binaries']) == sorted(VARIANTS)
    assert [r['name'] for r in build['rows']] == VARIANTS
    for row in build['rows']:
        assert row['exit_code'] == 0
        artifact(row['log'])
    assert all(a['ended'] <= b['started'] for a,b in zip(build['rows'],build['rows'][1:]))
    target = Path(origin['target'])
    if target.exists():
        for a in build['binaries'].values(): assert c.artifact(a['path']) == a
    else:
        cleanup = c.read(P/'cleanup.json')
        assert set(cleanup) == {'target','target_removed','removed_binaries','removed_target_bytes'}
        assert cleanup['target'] == str(target) and cleanup['target_removed'] is True
        assert cleanup['removed_binaries'] == list(build['binaries'].values())
        assert cleanup['removed_target_bytes'] >= sum(a['bytes'] for a in build['binaries'].values())
    quality = receipts('quality',3,'gates')
    assert 'fmt' in quality[0]['command'] and '--check' in quality[0]['command']
    for row, n in zip(quality[1:], [28,7]):
        assert 'test' in row['command'] and '--locked' in row['command']
        assert f'test result: ok. {n} passed; 0 failed;' in artifact(row['log']).read_text()
    oracle = {shape: inherited(f'qualification/0-{shape}-capture-before.json') for shape in SHAPES}
    reports, samples = 0, 0
    observed = {}
    lanes = {}
    for lane,count in [('controls',40),('heaptrack',10)]:
        lane_rows = receipts(lane,count,'children'); lanes[lane] = lane_rows
        expected = ([(r,s,v) for r,order in enumerate(plan[lane]['orders']) for s in SHAPES for v in order]
                    if lane == 'controls' else [(r,s,'profile') for r,order in enumerate(plan[lane]['orders']) for s in order])
        assert [(r['repeat'],r['shape'],r['variant']) for r in lane_rows] == expected
        for row in lane_rows:
            r,s,v = row['repeat'],row['shape'],row['variant']
            assert row['binary'] == build['binaries'][v]
            raw = c.read(artifact(row['report'])); ref = oracle[s]
            for key in ['schema','tool','mode','shape','slides','shapes_per_slide','timing_scope','marker','source','fixture']:
                assert raw[key] == ref[key], (lane,s,key)
            n = 3 if lane == 'controls' else 1
            assert len(raw['samples']) == raw['samples_requested'] == n and raw['warmup'] == 0
            cmd = ['taskset','-c','12']
            if lane == 'heaptrack':
                assert len(row['traces']) == 1
                artifact(row['traces'][0])
                cmd += ['heaptrack','--record-only','-o',row['report']['path'].removesuffix('.json')+'.heaptrack']
            cmd += [row['binary']['path'],'--mode','capture','--shape',s,'--samples',str(n),'--warmup','0','--output',row['report']['path']]
            assert row['command'] == cmd
            for i,sample in enumerate(raw['samples']):
                assert sample['index'] == i and sample['elapsed_ns'] > 0
                assert sample['source_sha256'] == ref['source']['sha256']
                assert sample['output'] == ref['samples'][0]['output']
                assert sample['verification'] == ref['samples'][0]['verification']
                assert {k:v for k,v in sample['metrics'].items() if k!='elapsed_ns'} == {k:v for k,v in ref['samples'][0]['metrics'].items() if k!='elapsed_ns'}
                if 'allocation' in v:
                    a = counter(sample['allocation'])
                    observed.setdefault(s,a); assert observed[s] == a
                else:
                    assert 'allocation' not in sample
            reports += 1; samples += n
    assert build['rows'][-1]['ended'] <= quality[0]['started']
    assert quality[-1]['ended'] <= lanes['controls'][0]['started']
    assert lanes['controls'][-1]['ended'] <= lanes['heaptrack'][0]['started']
    decodes = receipts('decoded',20,'decodes')
    assert lanes['heaptrack'][-1]['ended'] <= decodes[0]['started']
    rows = []
    for i,capture in enumerate(lanes['heaptrack']):
        decoded = {}
        for row,scope in zip(decodes[2*i:2*i+2],['whole','owner']):
            assert (row['repeat'],row['shape'],row['scope']) == (capture['repeat'],capture['shape'],scope)
            assert row['trace'] == capture['traces'][0]
            artifact(row['trace']); assert artifact(row['stderr']).stat().st_size == 0
            cmd = ['heaptrack_print','-f',row['trace']['path'],'-m','0','-t','0','-p','1','-a','1','-T','0','-n','20','--flamegraph-cost-type','allocations','-F',row['stacks']['path'],'-H',row['histogram']['path']]
            if scope == 'owner': cmd += ['--filter-bt-function','capture_region_0793']
            assert row['command'] == cmd
            stacks = parse_stacks(artifact(row['stacks']))
            hist = [tuple(map(int,line.split())) for line in artifact(row['histogram']).read_text().splitlines()]
            assert all(len(x)==2 and x[0]>=0 and x[1]>0 for x in hist)
            summary = int(re.search(r'calls to allocation functions:\s*(\d+)',artifact(row['log']).read_text())[1])
            decoded[scope] = (stacks, sum(n for _,n in hist), summary)
        whole,owner = decoded['whole'][0],decoded['owner'][0]
        total = sum(whole.values()); count = sum(owner.values())
        assert total == decoded['whole'][1] == decoded['whole'][2] == decoded['owner'][1] == decoded['owner'][2]
        selected = {s:n for s,n in whole.items() if any(OWNER.fullmatch(f) for f in s.split(';'))}
        assert owner == selected and count>0
        assert all('opened_presentation' in s for s in owner)
        a = observed[capture['shape']]
        expected = a['allocation_calls'] + a['reallocation_calls']
        rows.append({'repeat':capture['repeat'],'shape':capture['shape'],'whole_calls':total,'owner_calls':count,
                     'frozen_expected':expected,'frozen_pass':count==expected,
                     'supplementary_counter_match':count==a['allocation_calls'],
                     'nested_duplicate_checks':sum(n for s,n in owner.items() if 'check_for_duplicates' in s),
                     'nested_notes_inspector':sum(n for s,n in owner.items() if 'notes::codec::inspect_element' in s)})
    assert (reports,samples)==(50,130)
    assert all(not r['frozen_pass'] and r['supplementary_counter_match'] for r in rows)
    return {'schema':'litchi.performance.0793.analysis.v1','reports':reports,'samples':samples,
            'production_files':len(source['files']),'architecture_inputs':len(inputs),
            'fresh_probe_tests':35,'builds':4,'decodes':20,'operation_counters':observed,
            'rows':rows,'frozen_qualification':'fail','owner_fractions_authorized':False,
            'production_change':False,'speedup_claim':False}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--write',action='store_true'); parser.add_argument('--check',action='store_true')
    args = parser.parse_args(); result=analyze()
    if args.write: c.write(P/'analysis.json',result)
    if args.check: assert c.read(P/'analysis.json') == result
    print('0793 replay PASS: 50 reports / 130 samples; frozen qualification FAIL (10/10)')

if __name__=='__main__': main()
