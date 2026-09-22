"""Separate raw-period replay, custody checks and parser context controls."""
import collections,gzip,hashlib,importlib.util,json,re,statistics
from contract import P,ROOT,read,sha,guard,projection

def replay(text):
    buckets={};stages={};n=total=unknown=depth=0
    prefix='litchi_pptx::opened::cross_copy_plan::'
    root='litchi_perf_baseline::run_pptx_cross_copy_lifecycle'
    plan=prefix+'plan_cross_slide_copy_for_slides';prep=prefix+'prepare_cross_slide_copy_for_slides'
    applies={prefix+'apply_plan','litchi_pptx::package::model::Package::apply_cross_slide_copy_plan'}
    for block in text.strip().split('\n\n'):
        lines=block.strip().splitlines();header=lines.pop(0)
        assert header.rstrip().endswith('cycles:u:')
        period=int(header.split()[-2]);assert period>=0
        frames=[re.fullmatch(r'\s*[0-9a-f]+ (.*) \(.*\)',l).group(1) for l in lines]
        n+=1;total+=period;bad=any('[unknown]' in f for f in frames);unknown+=bad;depth+=len(frames)>=127
        has_apply=bool(applies.intersection(frames))
        if not frames or bad or len(frames)>=127 or (plan in frames and has_apply):c='ambiguous'
        elif root not in frames:c='unrooted_setup_oracle'
        elif has_apply:c='rooted_apply_excluded'
        elif plan in frames and prep in frames:
            c='rooted_plan_prepare' if frames.index(prep)<frames.index(plan)<frames.index(root) else 'ambiguous'
        elif prep in frames:c='ambiguous'
        else:c='rooted_other'
        b=buckets.setdefault(c,{'samples':0,'period':0});b['samples']+=1;b['period']+=period
        if c=='rooted_plan_prepare':
            s='plan_other'
            if prefix+'build_candidate' in frames:
                s='candidate_other'
                if 'litchi_opc::pkgwriter::PackageWriter::write_to_stream' in frames:
                    s='candidate_writer_other'
                    if 'soapberry_zip::preserve::generated_entry' in frames and 'zlib_rs::deflate::deflate' in frames:s='candidate_writer_generated_deflate'
            b=stages.setdefault(s,{'samples':0,'period':0});b['samples']+=1;b['period']+=period
    return {'buckets':buckets,'strict_plan_partition':stages,'total_samples':n,'total_period':total,'unknown_chain_samples':unknown,'depth_limit_samples':depth}

def main():
    guard();f=read('freeze.json');assert all(sha(P/n)==h for n,h in f['scripts'].items());assert sha(P/'build-fp.json')==f['build_sha256']
    archives={r['raw']:r for r in read('raw-archives.json')}
    for r in archives.values():
        assert sha(P/r['archive'])==r['archive_sha256'];h=hashlib.sha256();size=0
        with gzip.open(P/r['archive'],'rb') as stream:
            while chunk:=stream.read(1024*1024):h.update(chunk);size+=len(chunk)
        assert h.hexdigest()==r['raw_sha256'] and size==r['raw_bytes']
    receipts=read('formal/manifest.json');analysis=read('analysis.json');assert len(receipts)==len(analysis['rows'])==6
    cases=['pptx_cross_copy_plain_lifecycle','pptx_cross_copy_media_rich_lifecycle']
    expected=[cases[0],cases[1],cases[1],cases[0],cases[0],cases[1]]
    ratios=collections.defaultdict(list)
    for i,(r,a) in enumerate(zip(receipts,analysis['rows'])):
        assert r['mode']=='profile' and r['case']==expected[i] and r['samples']==(100 if 'plain' in r['case'] else 10)
        assert r['exit']==r['script_exit']==0
        for name,h in r['files'].items():
            path=f'formal/{name}'
            assert (archives[path]['raw_sha256'] if path in archives else sha(P/path))==h,path
        row={'case':r['case'],'samples':r['samples'],'warmups':0,'lane':'native','build':'fp'}
        assert projection(read(f'formal/{i:02d}.json'),row)==read('oracle.json')[r['case']]
        got=replay((P/f'formal/{i:02d}.stacks').read_text())
        for k,v in got.items():assert a[k]==v,(i,k)
        plan=got['buckets']['rooted_plan_prepare']['period']
        assert sum(v['period'] for v in got['strict_plan_partition'].values())==plan
        assert sum(v['period'] for v in got['buckets'].values())==got['total_period']
        ratio=100*got['strict_plan_partition'].get('candidate_writer_generated_deflate',{'period':0})['period']/plan
        assert abs(ratio-a['generated_deflate_pct_strict_plan_period'])<1e-12;ratios[r['case']].append(ratio)
    for c,values in ratios.items():assert analysis['summary'][c]['generated_deflate_pct_strict_plan_period']=={'min':min(values),'median':statistics.median(values),'max':max(values)}
    spec=importlib.util.spec_from_file_location('frozen',P/'parse-stacks.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
    base=[m.PREP,m.PLAN,m.ROOT]
    controls=[('strict',base,'rooted_plan_prepare'),('missing_root',base[:-1],'unrooted_setup_oracle'),('missing_plan',[m.PREP,m.ROOT],'ambiguous'),('apply_conflict',[m.APPLY[0]]+base,'ambiguous'),('public_apply_conflict',[m.APPLY[1]]+base,'ambiguous'),('root_only',[m.ROOT],'rooted_other'),('unknown',['[unknown]']+base,'ambiguous'),('empty',[],'ambiguous'),('reversed',base[::-1],'ambiguous'),('depth_limit',['leaf']*124+base,'ambiguous')]
    for name,frames,want in controls:assert m.category(frames)==want,name
    rejected=[]
    for name,text in [('bad_header','broken\n\t123 leaf (elf)'),('bad_frame','x 1/1 1.0: 3 cycles:u:\nwrong'),('lost','PERF_RECORD_LOST')]:
        try:list(m.parse(text))
        except (AssertionError,ValueError):rejected.append(name)
        else:raise AssertionError(name)
    return {'status':'passed','formal_processes':6,'archived_raw_captures':len(archives),'context_controls':[x[0] for x in controls],'malformed_controls':rejected,'scope':'separate replay of all formal bucket/stage periods and diagnostic percentages; exact frozen scripts, report and raw artifact custody'}
if __name__=='__main__':
    d=main();(P/'audit-root.json').write_text(json.dumps(d,indent=2)+'\n');print(json.dumps(d,indent=2))
