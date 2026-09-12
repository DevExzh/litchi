"""Validate the diagnostic native/profile outputs and retain same-build variation."""
import argparse
import importlib.util
import json
import math
from pathlib import Path
import run as R

HERE=R.HERE
FOLDER=HERE/'baseline'
PRIOR=HERE.parent/'change-0534'/'analyze.py'
spec=importlib.util.spec_from_file_location('numeric_0534_for_0535',PRIOR)
A=importlib.util.module_from_spec(spec)
spec.loader.exec_module(A)
# The prior row/catalog validators intentionally use this campaign's driver.
# Its old matrix/admission functions are never called.
assert A.RUN is R and A.HERE == HERE


def read(path):return json.loads(path.read_text())


def validate_job(lane,job):
    name=job['name'];path=FOLDER/(name+'.json')
    receipt=read(FOLDER/(name+'.receipt.json'))
    assert receipt['exit_code']==0 and receipt['script_sha256']==R.sha(HERE/'run.py')
    assert receipt['plan_sha256']==R.sha(HERE/'plan.json')
    assert receipt['source_manifest_sha256']==R.sha(FOLDER/'source-manifest.json')
    assert receipt['execution_stage']=='baseline' and receipt['execution_manifest_sha256']==receipt['source_manifest_sha256']
    meta=read(FOLDER/'binary-normal.json')
    assert receipt['binary_sha256']==meta['sha256']
    for name,digest in receipt['artifacts'].items():assert R.sha(FOLDER/name)==digest,name
    rows=A.validate_report(path,job)
    assert len(rows)==1
    return rows[0]


def analyze():
    result=[];identities={};profiles=[]
    for lane in ['native','profile']:
        for job in R.jobs(lane):
            row=validate_job(lane,job)
            key=(row['case'],row['corpus']['shape'])
            ident=A.identity(row)
            if key in identities:assert identities[key]==ident
            else:identities[key]=ident
            base=dict(name=job['name'],repeat=job['repeat'],case=key[0],shape=key[1],
                samples=job['samples'],identity=ident,report_sha256=R.sha(FOLDER/(job['name']+'.json')),
                receipt_sha256=R.sha(FOLDER/(job['name']+'.receipt.json')))
            if lane=='native':
                rss=read(FOLDER/(job['name']+'.rss.json'))
                assert set(rss)=={'max_rss_kib','elapsed_seconds','user_seconds','system_seconds'}
                assert all(isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v) and v>=0 for v in rss.values())
                base.update(timing_stats=A._elapsed_summary(row['elapsed_ns']),rss=rss)
                result.append(base)
            else:profiles.append(base)
    flags=[]
    for key in identities:
        pair=sorted([r for r in result if (r['case'],r['shape'])==key],key=lambda r:r['repeat'])
        assert len(pair)==2
        left,right=pair
        for group in ['timing_stats','rss']:
            for metric,a in A._numeric_leaves(left[group]):
                b=dict(A._numeric_leaves(right[group]))[metric]
                pct=100*(b/a-1) if a else (0 if b==0 else None)
                if pct is None or abs(pct)>5:
                    flags.append(dict(case=key[0],shape=key[1],metric=group+'.'+metric,
                        first=a,second=b,change_percent=pct,repeat_first=1,repeat_second=2,
                        review='Retain both same-build observations. No causal attribution, sample removal, stable-tail or before/after speedup claim follows.'))
    return dict(status='pass',schema='litchi-0535-diagnostic-native-v1',
        plan_sha256=R.sha(HERE/'plan.json'),source_manifest_sha256=R.sha(FOLDER/'source-manifest.json'),
        binary_sha256=read(FOLDER/'binary-normal.json')['sha256'],
        helper_sha256={str(PRIOR.relative_to(R.REPO)):R.sha(PRIOR),str(A.HELPER.relative_to(R.REPO)):R.sha(A.HELPER)},
        native_samples=sum(r['samples'] for r in result),profile_timing_samples_excluded=sum(r['samples'] for r in profiles),
        native_rows=result,profile_rows=profiles,same_build_variations_over_five_percent=flags,
        scope='Two warm synthetic in-memory baseline scenarios; profile clocks excluded from native latency; normal allocator metrics unavailable; no candidate, allocation or speedup claim')


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('output',nargs='?',type=Path)
    args=parser.parse_args();destination=args.output or HERE/'analysis.json'
    destination.write_text(json.dumps(analyze(),indent=2)+'\n')
    print('Native/profile report identity and baseline variation verified')
