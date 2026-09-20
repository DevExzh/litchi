#!/usr/bin/env python3
"""Replay paired owner attribution and separate process RSS diagnostics."""
import argparse,importlib.util,json,re
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('analysis0715mechanism',P/'analyze.py');A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)
C=A.C;M=A.module(P/'capture-mechanism.py','capture0715mechanismanalysis')

def analyze():
    plan=C.read(P/'plan.json');candidate=C.read(P/'source-candidate.json');assert C.C.census()==candidate
    pilot=C.read(P/'pilot-analysis.json');assert pilot['decision']['accepted']
    for n,h in C.read(P/'mechanism-freeze.json').items():assert C.C.sha(P/n)==h
    base=C.read(P/'profile-analysis.json');T=A.module(P/'analyze_profiles.py','parser0715mechanism')
    rows=[]
    for job in M.jobs(plan):
        name=job['name'];r=C.read(P/(name+'.receipt.json'));b=M.binary(job)
        assert r['job']==job and r['command']==M.command(plan,job,b) and r['exit_code']==0 and r['binary']==b
        binary=Path(b['path'])
        if binary.exists():assert C.C.sha(binary)==b['sha256'] and binary.stat().st_size==b['bytes']
        else:assert b in C.read(P/'cleanup.json')['binaries']
        assert r['live_source_sha256']==C.C.sha(P/'source-candidate.json')
        assert r['binary_source_sha256']==C.C.sha(P/('source-'+job['source']+'.json'))
        assert r['script_sha256']==C.C.sha(P/'capture-mechanism.py') and r['plan_sha256']==C.C.sha(P/'plan.json')
        assert r['pilot_analysis_sha256']==C.C.sha(P/'pilot-analysis.json') and r['fixture']==C.fixture(job['corpus'])
        expected={name+s for s in ['.json','.stdout','.stderr']}
        if job['kind']=='rss':expected.add(name+'.time-v')
        else:
            expected.add(name+'.callgrind');expected.update(name+'.callgrind.'+str(i) for i in range(1,plan['profile']['numbered_parts'][job['corpus']['id']]+1))
        assert set(r['artifacts'])==expected
        for n,h in r['artifacts'].items():assert C.C.sha(P/n)==h
        report=C.read(P/(name+'.json'));meta=dict(binary_sha256=b['sha256'],binary_bytes=b['bytes'])
        A.L.check_report_metadata(report,meta,'native',job['case'],job['samples'],job['warmup'],name)
        result=report['results'][0];A.L.validate_elapsed(result['elapsed_ns'],job['samples'],name)
        A.L.validate_operation_metrics(result['operation_metrics'],result['elapsed_ns'],job['samples'],'native',name)
        A.L.validate_ordinary_save(result,job['corpus'],job['phase'],'native',job['samples'],name)
        reference=C.read(P.parent/'change-0714'/('native-r1-'+job['corpus']['id']+'-'+job['phase']+'.json'))['results'][0]
        assert A.N.normalized_result(result)==A.N.normalized_result(reference)
        row=dict(name=name,source=job['source'],kind=job['kind'],repeat=job['repeat'],corpus=job['corpus']['id'],phase=job['phase'],receipt_sha256=C.C.sha(P/(name+'.receipt.json')))
        if job['kind']=='profile':
            row['profile']=T.analyze_profile(P/(name+'.callgrind'),job['corpus']['id'],plan['profile'])
            baseline=next(x['profile'] for x in base['profiles'] if x['corpus']==job['corpus']['id'] and x['repeat']==job['repeat'])
            row['baseline_owner_ir']=baseline['owner_ir'];row['candidate_owner_ir']=row['profile']['owner_ir'];row['delta_percent']=(row['candidate_owner_ir']/row['baseline_owner_ir']-1)*100
        else:
            raw=(P/(name+'.time-v')).read_text();rss=re.findall(r'^\s*Maximum resident set size \(kbytes\): (\d+)\s*$',raw,re.M);status=re.findall(r'^\s*Exit status: (\d+)\s*$',raw,re.M)
            assert len(rss)==1 and int(rss[0])>0 and status==['0'];row['max_rss_kib']=int(rss[0])
        rows.append(row)
    pairs=[];repeats=[]
    for corpus in ['generated','numbered-list']:
        for phase in ['counting_publish','lifecycle']:
            for repeat in [1,2]:
                chosen={r['source']:r['max_rss_kib'] for r in rows if r['kind']=='rss' and r['corpus']==corpus and r['phase']==phase and r['repeat']==repeat}
                assert set(chosen)=={'baseline','candidate'};delta=(chosen['candidate']/chosen['baseline']-1)*100
                pairs.append(dict(corpus=corpus,phase=phase,repeat=repeat,**chosen,delta_percent=delta,flag=delta>5))
            for source in ['baseline','candidate']:
                values=[r['max_rss_kib'] for r in rows if r['kind']=='rss' and r['corpus']==corpus and r['phase']==phase and r['source']==source];assert len(values)==2
                spread=(max(values)/min(values)-1)*100;repeats.append(dict(corpus=corpus,phase=phase,source=source,values=values,spread_percent=spread,flag=spread>5))
    return dict(status='pass',children=rows,profile_children=4,rss_children=16,normalized_parity=True,rss_pairs=pairs,rss_repeat_review=repeats,limitations=C.read(P/'mechanism-plan.json')['claims'])

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args();out=P/'mechanism-analysis.json';data=(json.dumps(analyze(),indent=2,sort_keys=True)+'\n').encode()
    if args.check:assert out.read_bytes()==data
    else:assert not out.exists();out.write_bytes(data)
    print('PASS: four paired owner profiles, 16 RSS diagnostics and exact parity')
if __name__=='__main__':main()
