#!/usr/bin/env python3
"""Attribute sealed 0448 samples; no workload, profiler, Git or binary required."""
import argparse, collections, gzip, hashlib, importlib.util, json, re
from pathlib import Path
ROOT = Path(__file__).resolve().parent
HEADER = re.compile(r'^\S.*?\s+\d+\s+[0-9.]+:\s+(\d+)\s+cycles:u:\s*$')
FRAME = re.compile(r'^\s*[0-9a-f]+\s+(.+?)\s+\([^\n]*\)\s*$')
RUN = 'litchi_perf_baseline::pptx_provider_lifecycle::run_lifecycle_iteration'
def require(ok, label):
    if not ok: raise ValueError(label)
def load(name): return json.loads((ROOT/name).read_text())
def digest(raw): return hashlib.sha256(raw).hexdigest()
def parse(data):
    result=[];period=None;frames=[]
    def finish():
        if period is not None: result.append((period,tuple(frames)))
    for line in data.decode().splitlines():
        if match:=HEADER.fullmatch(line):
            finish();period=int(match[1]);frames=[]
        elif not line.strip():
            finish();period=None;frames=[]
        elif period is not None and (match:=FRAME.fullmatch(line)):
            frames.append(re.sub(r'\+0x[0-9a-f]+$','',match[1]))
        else: raise ValueError('unparsed profile line: '+line)
    finish();require(bool(result),'empty profile');return result

def category(frames):
    if not any(RUN in f for f in frames): return 'outside-lifecycle-frame'
    if any('litchi_perf_baseline::sha256_hex' == f for f in frames): return 'untimed-harness-output-hash'
    if any('::digest_touched' in f for f in frames):
        if any('::publish_cross_slide_copy_to_stream' in f for f in frames): return 'publication-touched-digest'
        if any('::plan_cross_slide_copy' in f for f in frames): return 'planning-touched-digest'
    return 'unclassified-lifecycle-sha'

def profile(lane):
    receipt=load(f'inputs/profiles/{lane}/record/receipt.json')
    record=receipt['artifacts']['perf_script']
    raw=gzip.decompress((ROOT/'inputs'/Path(record['path']+'.gz')).read_bytes())
    require(digest(raw)==record['sha256'] and len(raw)==record['bytes'],'raw stack identity')
    blocks=parse(raw);whole=sum(p for p,_ in blocks)
    run=sum(p for p,f in blocks if any(RUN in v for v in f))
    sha=collections.Counter();counts=collections.Counter();stacks=collections.Counter()
    for period,frames in blocks:
        if frames and frames[0].startswith('sha2::sha256::'):
            c=category(frames);sha[c]+=period;counts[c]+=1;stacks[(c,frames)]+=period
    total_sha=sum(sha.values());run_sha=total_sha-sha['outside-lifecycle-frame']
    original=load('inputs/profile-summary.json')['profiles'][0 if lane==1 else 1]
    require((len(blocks),whole,run)==(original['sample_blocks'],original['total_period'],original['run_frame_period']),'original period/count conservation')
    require(run_sha==sum(r['period'] for r in original['run_frame_self'] if r['symbol'].startswith('sha2::sha256::')),'original SHA conservation')
    require(total_sha==sum(r['period'] for r in original['whole_process_self'] if r['symbol'].startswith('sha2::sha256::')),'whole SHA conservation')
    require(sum(stacks.values())==total_sha,'caller stack conservation')
    rows=[]
    for c,p in sorted(sha.items()):
        rows.append({'category':c,'blocks':counts[c],'period':p,'percent_of_all_sha':100*p/total_sha,'percent_of_run_period':None if c=='outside-lifecycle-frame' else 100*p/run,'percent_of_run_sha':None if c=='outside-lifecycle-frame' else 100*p/run_sha})
    return {'lane':lane,'policy':receipt['lane']['minimum_service'] and 'minimum-service' or 'separate-sleeps','sample_blocks':len(blocks),'total_period':whole,'run_period':run,'sha_period':total_sha,'run_sha_period':run_sha,'missing_callchain_blocks':sum(not f for _,f in blocks),'missing_callchain_period':sum(p for p,f in blocks if not f),'categories':rows,'sha_caller_stacks':[{'category':c,'frames':list(f),'period':p} for (c,f),p in sorted(stacks.items(),key=lambda r:(-r[1],r[0]))]}

def reports():
    spec=importlib.util.spec_from_file_location('oracle',ROOT/'inputs/verify-report.py')
    oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
    by_corpus={};samples=0
    for i in range(8):
        raw=(ROOT/f'inputs/runs/{i}/report.json').read_bytes();report=json.loads(raw)
        receipt=load(f'inputs/runs/{i}/receipt.json');record=receipt['artifacts']['report']
        require(digest(raw)==record['sha256'] and len(raw)==record['bytes'],'report identity')
        oracle.check_report(report);require((report['samples'],report['warmup'])==(30,3),'sample dimensions')
        for sample in report['samples_raw']:
            rows=[]
            for phase in sample['phases']:
                if phase['label'] not in ['opened','planned','published']:continue
                row={'phase':phase['label']}
                for owner in ['source','destination']:
                    d=phase[owner+'_reads']['delta'];require(d is not None,'phase delta')
                    row[owner+'_reads']={k:d[k] for k in ['logical_calls','returned_bytes','short_reads','request_size_counts','delayed_calls','transfer_paced_calls','transfer_delay_ns']}
                    c=phase[owner+'_cache']
                    row[owner+'_cache']={k:c[k] for k in ['availability','event_delta','retained_entries','retained_bytes']}
                row['combined_returned_bytes']=sum(row[o+'_reads']['returned_bytes'] for o in ['source','destination'])
                rows.append(row)
            require(len(rows)==3,'phase coverage')
            corpus=report['corpus']
            if corpus in by_corpus:require(by_corpus[corpus]==rows,'work/cache identity across every sample and policy')
            else:by_corpus[corpus]=rows
            samples+=1
    require(samples==240,'retained samples')
    return {'reports':8,'samples':samples,'by_corpus':by_corpus}

def derive():
    inputs=load('inputs.json')
    for r in inputs['records']:
        path=Path(r['path']);require(not path.is_absolute() and '..' not in path.parts,'safe input path')
        data=(ROOT/path).read_bytes();require(len(data)==r['bytes'] and digest(data)==r['sha256'],'input custody '+str(path))
    return {'change':449,'profiles':[profile(i) for i in [1,3]],'work':reports(),'scope':'reanalysis of 0448; caller-stack subsets include warmups, not exact timer intervals; CPU periods omit blocked sleeps; per-owner source totals are not per-member offset traces'}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();result=derive()
    if a.check:require(load('attribution.json')==result,'derived attribution')
    else:
        with (ROOT/'attribution.json').open('x') as f:f.write(json.dumps(result,indent=2)+'\n')
    print(json.dumps({'status':'pass','reports':8,'samples':240,'profiles':2,'check':a.check}))
