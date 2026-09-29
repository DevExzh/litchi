"""Independent failure-packet audit. Does not import the admission reader."""
from pathlib import Path
import sys
import driver as d
import reuse_quality

P=d.P
OPC='a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6'
PPTX='61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757'
OUTPUT='f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009'
ALIGNED_OUTPUT='3f66abdd5fbdc94eaf4961089cd7193deb9eaffff9004871e8f4e47544a55a66'
PPTX_ERROR='PPTX source replay violated pptx_file_source_open_selected_slide_lifecycle payload-range classification'
OPC_ERROR='eager and source-backed OPC filesystem save samples differ'

def descriptor(record): assert d.desc(record['path'])==record

def proof(p,metrics):
    assert p['status']=='eligible' and p['fsync_completed'] is True
    assert p['advice']=='posix_fadvise_dontneed_accepted'
    assert p['filesystem_magic']==0xef53
    assert p['source_bytes']==p['aligned_source_bytes']==p['fincore_size_bytes']
    assert p['source_bytes']%p['page_size_bytes']==0
    assert p['source_pages']*p['page_size_bytes']==p['source_bytes']
    assert all(p[k]==0 for k in ('resident_bytes','dirty_bytes','writeback_bytes'))
    assert p['read_bytes_after']-p['read_bytes_before']==p['read_bytes_delta']>0
    assert p['read_bytes_delta']==metrics['read_bytes']
    q=p['fincore_post']
    assert q['status']=='eligible' and q['size_bytes']==p['source_bytes']
    assert 0<q['resident_bytes']<=q['size_bytes']
    assert q['dirty_bytes']==q['writeback_bytes']==0
    for k in ['fincore_tool','fincore_sha256','fincore_version','fincore_method','fincore_fallback']:
        assert q[k]==p[k]
    assert p['fincore_sha256']==d.read(P/'prepare.json')['fincore']['sha256']
    for v in (p,q):
        assert v['fincore_stderr_bytes']==v['fincore_version_stderr_bytes']==0
        assert v['fincore_fallback']=='none'

def report(path,case,states):
    r=d.read(path);b=d.read(P/'build.json')['binary']
    assert r['schema_version']==1
    assert r['binary_identity']['binary_sha256']==b['sha256']
    assert r['binary_identity']['binary_bytes']==b['bytes']
    c=r['configuration']
    assert c['cases']==[case] and c['filesystem_cache_states']==states
    assert c['samples_per_case']==1 and c['warmup_iterations_per_case']==0
    assert c['filesystem_fresh_child_per_sample'] and c['filesystem_process_isolated'] and c['filesystem_root_selected']
    assert len(r['filesystem_evidence'])==1
    e=r['filesystem_evidence'][0]
    assert e['case']==case and e['cache_states']==states and e['sample_count']==1
    assert e['fresh_child_per_sample'] and e['warmup_iterations']==0
    assert e['corpus']['archive_sha256']==(OPC if case.startswith('opc') else PPTX)
    assert [s['cache_state'] for s in e['samples']]==states
    assert len(r['results'])==len(states)
    for s in e['samples']:
        assert s['sample_index']==0 and s['child_process_id']>0
        assert 0<s['elapsed_ns']<=s['parent_wall_ns']
        sizes=s['logical_read_request_sizes']
        assert len(sizes)==s['logical_read_calls']
        assert sum(sizes)==s['logical_read_requested_bytes']>=s['logical_read_bytes']
        assert sum(s['logical_read_request_size_buckets'].values())==len(sizes)
        result=next(x for x in r['results'] if x['cache_state']==s['cache_state'])
        assert result['case']==case and result['corpus']==e['corpus']
        assert result['elapsed_ns']['samples']==[s['elapsed_ns']]
        assert result['elapsed_ns']['sample_order']==[0]
        if s['cache_state']=='cold-verified':
            assert e['cold_verified_status']=='eligible'
            assert e['cold_verified_samples']==[s['cold_verified']]
            proof(s['cold_verified'],s['process_metrics'])
        else: assert s.get('cold_verified') is None
        if case.startswith('opc'):
            assert s['opc_materialized_parts']==(4 if '_eager_' in case else 0)
            assert s['logical_read_counter_scope']==('not_applicable_eager_opc' if '_eager_' in case else 'timed_read_at')
        if 'save' in case:
            aligned=case=='opc_file_source_one_part_atomic_save' and s['cache_state']=='cold-verified'
            assert s['output_sha256']==(ALIGNED_OUTPUT if aligned else OUTPUT)
            assert s['output_bytes']==(16785408 if aligned else 16783632)
            assert result['output_sha256']==s['output_sha256']
        else: assert s['output_sha256'] is None and s['output_bytes'] is None
        if case=='pptx_file_source_open_selected_slide_lifecycle':
            assert s['logical_read_counter_scope']=='untimed_source_replay_only'
            p=s['pptx_source_replay']
            assert p['source_sha256']==PPTX and p['selected_position']==100 and p['slide_count']==200
            assert p['selected_slide_payload_fully_covered']
            assert p['selected_slide_payload_covered_bytes']==522
            assert p['unselected_slide_payload_read_bytes']==p['media_payload_read_bytes']==0
            assert p['semantic_sha256']=='f5f7db181150c00a4323a48c142721ead73aca3ad7c3b3594e8b1a18a686b257'
    return len(e['samples'])

def derive():
    p=d.read(P/'prepare.json'); assert d.inventory()==p['source']
    assert all(d.sha(d.ROOT/k)==v for k,v in p['unrelated'].items())
    assert all(d.sha(P/k)==v for k,v in p['frozen'].items())
    assert d.read(P/'quality.json')==reuse_quality.value()
    build=d.read(P/'build.json');descriptor(build['receipt'])
    br=d.read(build['receipt']['path']);assert br['exit_code']==0 and br['error'] is None
    descriptor(br['log'])
    if d.BINARY.exists(): descriptor(build['binary'])
    else: assert d.read(P/'cleanup.json')['binary']==build['binary']
    q=d.read(P/'qualification.json');assert q['status']=='failed' and len(q['rows'])==6
    diag=d.read(P/'diagnostics.json');assert diag['performance_claim']=='none'
    observed=[]; total=0; last_end=br['finished_unix']
    specs=[(f'qualification-{i:02}',case,'warm,cold-verified',i!=5,
            None if i!=5 else PPTX_ERROR,q['rows'][i]) for i,case in enumerate(d.CASES)]
    ds=[('pptx-warm',d.CASES[5],'warm',True,None),
        ('pptx-cold',d.CASES[5],'cold-verified',False,PPTX_ERROR),
        ('opc-save-pair-cold',','.join(d.CASES[2:4]),'cold-verified',False,OPC_ERROR)]
    assert len(diag['results'])==3
    specs.extend((f'diagnostic-{label}',case,state,ok,error,row)
        for (label,case,state,ok,error),row in zip(ds,diag['results'],strict=True))
    for label,case,states,ok,error,row in specs:
        descriptor(row['receipt']); r=d.read(row['receipt']['path'])
        path=P/f'{label}.json'
        argv=['taskset','-c','12',str(d.BINARY),'--case',case,'--warmup','0','--samples','1',
              '--filesystem-cache',states,'--filesystem-root',str(d.SCRATCH),'--json',str(path)]
        assert r['argv']==argv
        assert r['started_unix']>=last_end and r['finished_unix']>=r['started_unix']
        last_end=r['finished_unix'];descriptor(r['log'])
        assert r['exit_code']==row['exit_code']==(0 if ok else 1) and r['error'] is None
        assert r['prepare_sha256']==d.sha(P/'prepare.json')
        if ok:
            descriptor(row['report']);assert row['report']['path']==str(path)
            total+=report(path,case,states.split(','))
        else:
            assert row['report'] is None and not path.exists()
            assert error in Path(r['log']['path']).read_text()
        observed.append(dict(label=label,exit_code=r['exit_code'],report_retained=ok))
    assert not list(P.glob('measured-*.json'))
    assert not (P/'capture.json').exists() and not (P/'capture-freeze.json').exists()
    assert total==11
    return dict(status='pass',disposition='qualification_failed',commands=observed,
      retained_reports=6,retained_qualification_diagnostic_samples=total,formal_reports=0,
      performance_claim='none',inputs_checked=len(p['source']),quality_gates_reused=8,
      opcode_cold_output_byte_difference=1776)

if __name__=='__main__':
    result=derive()
    if sys.argv[1:]==['--write']:d.write(P/'audit.json',result)
    else:
        assert not sys.argv[1:];assert d.read(P/'audit.json')==result
    print('Independent failure audit PASS: 6 reports, 11 diagnostic samples, zero formal reports')
