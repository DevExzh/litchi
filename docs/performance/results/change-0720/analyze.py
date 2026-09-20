#!/usr/bin/env python3
"""Replay the untimed diagnostic's exact-source and output checks."""
import json, sys
from custody import P, ROOT, TARGET, census, sha

def read(n): return json.loads((P/n).read_text())

def custody():
    source=read('baseline-build-source.json')
    assert source==census(), 'production source must be restored'
    for n in ['constraints.json','motivation.json']:
        for path,h in read(n).items(): assert sha(ROOT/path)==h,path
    binary=TARGET/'debug/docx-active-offset-oracle-0720'
    stages=['baseline-build','baseline','trace-build','trace','trace-repeat']
    for stage in stages:
        r=read(stage+'-receipt.json')
        assert r['stage']==stage and r['exit_code']==0
        assert r['cwd']==str(ROOT) and r['target']==str(TARGET) and r['cargo_build_jobs']==4
        expected=(['cargo','build','--locked','--manifest-path',str(P/'oracle/Cargo.toml')]
                  if stage.endswith('build') else [str(binary),'--output',str(P/(stage+'.json'))])
        assert r['argv']==expected
        assert r['source_sha256']==sha(P/(stage+'-source.json'))
        assert r['oracle_source_sha256']==sha(P/'oracle/src/main.rs')
        for suffix in ['stdout','stderr']: assert r[suffix+'_sha256']==sha(P/(stage+'.'+suffix))
        if not stage.endswith('build'): assert r['report_sha256']==sha(P/(stage+'.json'))
    assert read('baseline-source.json')==source
    assert read('baseline-build-receipt.json')['binary_sha256']==read('baseline-receipt.json')['binary_sha256']
    traced=read('trace-build-source.json')
    assert traced==read('trace-source.json')==read('trace-repeat-source.json')
    assert set(traced)==set(source)
    changed=[n for n in source if source[n]!=traced[n]]
    assert set(changed)=={'crates/litchi-docx/src/lib.rs','crates/litchi-docx/src/alt/codec.rs','crates/litchi-docx/src/namespace.rs','crates/litchi-docx/src/parts/document_part.rs','crates/litchi-docx/src/writer/doc/package.rs'},changed
    import instrument
    instrumentation=read('instrumentation.json')
    assert instrumentation['fragment_sha256']==sha(P/'trace.fragment')
    assert instrumentation['status']=='applied'
    assert set(instrument.SOURCE_SHA256)==set(changed)
    for row in instrumentation['sources']:
        name=row['relative']
        assert row['original_sha256']==source[name]==instrument.SOURCE_SHA256[name]
        assert row['transformed_sha256']==traced[name]
        assert instrument.digest(instrument.transform(name,(ROOT/name).read_bytes()))==traced[name]
    failed=read('initial-trace-build/trace-build-receipt.json')
    assert failed['exit_code']==101
    assert failed['stderr_sha256']==sha(P/'initial-trace-build/trace-build.stderr')
    assert read('initial-trace-build/instrumentation.json')['fragment_sha256']==sha(P/'initial-trace-build/trace.fragment')
    trace_binary=read('trace-build-receipt.json')['binary_sha256']
    assert all(read(n+'-receipt.json')['binary_sha256']==trace_binary for n in ['trace','trace-repeat'])
    assert (P/'baseline.json').read_bytes()==(P/'trace.json').read_bytes()==(P/'trace-repeat.json').read_bytes()
    assert (P/'trace.stderr').read_bytes()==(P/'trace-repeat.stderr').read_bytes()
    report=read('baseline.json');assert report['case_count']==19 and len(report['packages'])==2
    for row in report['packages']: assert row['outcome']=='Ok(alts=0)'
    quality=read('quality.json')
    expected=[('fmt',['cargo','fmt','--manifest-path',str(P/'oracle/Cargo.toml'),'--','--check']),
              ('clippy',['cargo','clippy','--locked','--manifest-path',str(P/'oracle/Cargo.toml'),'--all-targets','--','-D','warnings'])]
    expected += [(r['name'],r['command']) for r in read('../change-0717/evidence/results.json')]
    assert [(r['name'],r['argv']) for r in quality]==expected
    for row in quality:
        assert row['exit_code']==0 and row['log_sha256']==sha(P/('quality-'+row['name']+'.log'))
    if binary.exists(): assert sha(binary)==trace_binary
    else:
        cleanup=read('cleanup.json')
        assert cleanup['trace_binary']['sha256']==trace_binary
        assert cleanup['trace_binary']['path']==str(binary)
        assert cleanup['production_source_restored']
        from pathlib import Path
        assert all(not Path(n).exists() for n in cleanup['removed_paths'])
    if (P/'documentation.json').exists():
        for path,h in read('documentation.json').items():assert sha(ROOT/path)==h
    return changed

def traces(name):
    label=None;start=None;rows={};sequence=0
    for line in (P/name).read_text().splitlines():
        if line.startswith('CASE0720 '):
            assert start is None
            label=line.removeprefix('CASE0720 ')
            assert label not in rows;rows[label]=[]
        elif line.startswith('TRACE0720 '):
            row=json.loads(line.removeprefix('TRACE0720 '));assert label is not None
            if row['kind']=='document_start':
                assert start is None and row['sequence']==sequence+1
                sequence=row['sequence'];start=row
            else:
                assert row['kind']=='document_end' and start is not None
                for key in ['sequence','xml_bytes','xml_sha256']:assert row[key]==start[key]
                for record in row['alt_scans']+row['range_passes']+row['active_calls']:
                    assert record['xml_bytes']==row['xml_bytes'] and record['xml_sha256']==row['xml_sha256']
                rows[label].append(row);start=None
        else: raise AssertionError(('unexpected stderr',line))
    assert start is None
    return rows

def analyze():
    changed=custody()
    rows=traces('trace.stderr');assert rows==traces('trace-repeat.stderr')
    report=read('baseline.json')
    assert set(rows)=={'generated-setup','generated-medium','numbered-list',*(c['name'] for c in report['cases'])}
    summaries=[]
    import hashlib
    digest=lambda obj:hashlib.sha256(json.dumps(obj,sort_keys=True,separators=(',',':')).encode()).hexdigest()
    for name,records in rows.items():
        if name=='generated-setup':continue
        assert len(records)==1,(name,len(records))
        row=records[0];assert len(row['alt_scans'])==1
        controls={c['name']:c for c in report['cases']}
        if name in controls and name!='bom-plain':
            assert row['xml_bytes']==controls[name]['xml_bytes']
            assert row['xml_sha256']==controls[name]['xml_sha256']
        if name=='bom-plain': assert row['xml_bytes']==controls[name]['xml_bytes']-3
        if name=='numbered-list':
            import zipfile
            fixture=ROOT/'test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx'
            with zipfile.ZipFile(fixture) as archive: xml=archive.read('word/document.xml')
            assert row['xml_bytes']==len(xml) and row['xml_sha256']==hashlib.sha256(xml).hexdigest()
        alt=row['alt_scans'][0]
        assert len(row['range_passes'])<=1 and len(row['active_calls'])<=2
        if alt['completed']:
            assert len(row['range_passes'])==1
            assert alt['consumed_reader_bytes']==row['xml_bytes']
            assert row['active_calls'][0]['result_debug'].startswith('Ok(')
        else: assert not row['range_passes']
        for record in row['alt_scans']+row['range_passes']:
            assert 0<=record['consumed_reader_bytes']<=row['xml_bytes'] and record['event_count']>0
        for ranges in row['range_passes']:
            if ranges['raw_ranges'] is not None:
                if len(row['active_calls'])==2:
                    assert [r[1] for r in ranges['raw_ranges']]==row['active_calls'][1]['input_offsets']
                assert all(t in [0,1,2] and l>0 and s+l<=row['xml_bytes'] for t,s,l in ranges['raw_ranges'])
            if ranges['selected_ranges'] is not None:
                assert len(row['active_calls'])==2
                outcome=row['active_calls'][1]['result_debug']
                assert outcome.startswith('Ok(') and outcome.endswith(')')
                selected=json.loads(outcome[3:-1])
                assert ranges['selected_ranges']==[r for r in ranges['raw_ranges'] if r[1] in set(selected)]
        summary={'name':name,'xml_bytes':row['xml_bytes'],'xml_sha256':row['xml_sha256'],
            'alt_events':alt['event_count'],'alt_completed':alt['completed'],
            'range_events':[r['event_count'] for r in row['range_passes']],
            'range_completed':[r['completed'] for r in row['range_passes']],
            'mce_input_counts':[len(r['input_offsets']) for r in row['active_calls']],
            'mce_outcomes':[r['result_debug'] if r['result_debug'].startswith('Err(') else 'Ok' for r in row['active_calls']],
            'chunk_metadata_sha256':digest([alt['raw_chunks_debug'],alt['selected_chunks_debug']]),
            'range_metadata_sha256':digest([(r['raw_ranges'],r['selected_ranges']) for r in row['range_passes']]),
            'mce_inputs_outputs_sha256':digest(row['active_calls'])}
        if name in ['generated-medium','numbered-list']:
            assert len(row['range_passes'])==1 and len(row['active_calls'])==2
            ranges=row['range_passes'][0]
            assert alt['completed'] and ranges['completed']
            assert alt['event_count']==ranges['event_count']
            assert ranges['consumed_reader_bytes']==row['xml_bytes']
            assert row['active_calls'][0]['input_offsets']==[]
            summary['structural_passes']=2
            summary['sum_structural_source_bytes']=2*row['xml_bytes']
        summaries.append(summary)
    return {'status':'pass','public_report_parity':'byte-identical across baseline and two trace runs',
            'trace_repeat_parity':'byte-identical stderr','xml_controls':19,'package_corpora':2,
            'performance_claim':'none','temporary_instrumented_files':changed,'cases':summaries}

if __name__=='__main__':
    result=analyze()
    if sys.argv[1:]==['--write']:(P/'analysis.json').write_text(json.dumps(result,indent=2)+'\n')
    else:assert result==read('analysis.json')
    print('PASS: exact public outputs, repeated body traces, source and command custody')
