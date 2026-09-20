#!/usr/bin/env python3
"""Describe retained native sample order without inferring a cause or rerunning."""
import argparse,hashlib,json,statistics
from pathlib import Path
P=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def analyze():
    pilot=json.loads((P/'pilot-analysis.json').read_text());assert not pilot['decision']['accepted']
    rows={};series={}
    for stage in ['baseline-A1','candidate-B1','candidate-B2','baseline-A2']:
        name=stage.replace('-','_')+'-native-numbered-list-counting_publish';f=P/(name+'.json')
        bound=next(r for r in pilot['rows'] if r['stage']==stage and r['lane']=='native' and r['corpus_id']=='numbered-list' and r['phase']=='counting_publish');assert bound['report_sha256']==sha(f)
        elapsed=json.loads(f.read_text())['results'][0]['elapsed_ns'];values=elapsed['samples'];order=elapsed['sample_order'];assert len(values)==100 and sorted(order)==list(range(100))
        execution=[None]*100
        for value,index in zip(values,order,strict=True):execution[index]=value
        series[stage]=execution
        rows[stage]=dict(report_sha256=sha(f),min_ns=min(values),max_ns=max(values),p50_ns=elapsed['p50'],p95_ns=elapsed['p95'],p99_ns=elapsed['p99'],mean_ns=statistics.mean(values),blocks_of_25_mean_ns=[statistics.mean(execution[i:i+25]) for i in range(0,100,25)],half_mean_ns=[statistics.mean(execution[:50]),statistics.mean(execution[50:])])
    b1=series['candidate-B1'];a1=rows['baseline-A1'];counts={k:sum(v>a1[k] for v in b1) for k in ['p50_ns','p95_ns','p99_ns','max_ns']}
    return dict(pilot_analysis_sha256=sha(P/'pilot-analysis.json'),rows=rows,b1_samples_above_a1_thresholds=counts,decision='rejected; sustained B1 child shift retained',limits=['Execution order reconstructed from the retained sample_order permutation.','No hardware, allocator, frequency, scheduling or implementation cause identified.','No samples or children discarded; no rerun performed.'])
def main():
    parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args();data=(json.dumps(analyze(),indent=2,sort_keys=True)+'\n').encode();out=P/'child-shift.json'
    if args.check:assert out.read_bytes()==data
    else:assert not out.exists();out.write_bytes(data)
    print('PASS retained sample-order shift diagnostic')
if __name__=='__main__':main()
