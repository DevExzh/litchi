#!/usr/bin/env python3
"""Recompute allocator-policy diagnostics and ordinary process-fault observations."""
import csv,json,re,statistics
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def counters(path):
    result={}
    for row in csv.reader(path.read_text().splitlines()):
        if len(row)>4:
            assert float(row[4])==100.0
            result[row[2]]=int(row[0])
    assert len(result)==6
    return result
def medians(report):
    return {k:statistics.median(row['timings'][k] for row in report['samples_raw']) for k in report['samples_raw'][0]['timings']}
def derive():
    proof=json.loads((ROOT/'allocator-regime-proof.json').read_text());rows=[]
    for record in proof['rows']:
        path=ROOT/'allocator-regime'/str(record['lane'])
        rows.append({k:record[k] for k in ['lane','build','mmap_threshold']}|{'counters':counters(path/'perf.csv'),'median_ns':medians(json.loads((path/'report.json').read_text()))})
    pairs=[]
    for a,b in [(0,1),(3,2),(4,5),(7,6)]:
        pairs.append({'baseline_lane':a,'candidate_lane':b,'mmap_threshold':rows[a]['mmap_threshold'],'counter_percent':{k:100*(rows[b]['counters'][k]/v-1) for k,v in rows[a]['counters'].items()},'median_percent':{k:100*(rows[b]['median_ns'][k]/v-1) for k,v in rows[a]['median_ns'].items()}})
    ordinary=[]
    for base,indices in [(ROOT,[1,5,10,14])]:
        for i in indices:
            path=base/'runs'/str(i);r=json.loads((path/'receipt.json').read_text());report=json.loads((path/'report.json').read_text());resource=(path/'resource.log').read_text()
            ordinary.append({'path':str(path.relative_to(ROOT)),'build':r['lane_definition']['build'],'api_p50_ns':medians(report)['api_sum_ns'],'minor_faults':int(re.search(r'Minor \(reclaiming a frame\) page faults: (\d+)',resource)[1])})
    return {'scope':'Controlled whole-process perf stat diagnostic with fixed glibc mmap thresholds; not ordinary timing lanes or API-attributed counters. Original ordinary bytes/media timings remain adverse. Page-fault association is supported; exact historical allocation/map decisions were not traced.','rows':rows,'pairs':pairs,'ordinary_process_faults':ordinary}
def render(d):
    lines=['# Fixed allocator-policy diagnostic','',d['scope'],'','| Threshold | Pair | API p50 change | Instructions change | Cycles change | Page faults A / B |','|---:|---|---:|---:|---:|---:|']
    for p in d['pairs']:
        a=d['rows'][p['baseline_lane']];b=d['rows'][p['candidate_lane']]
        lines.append(f"| {p['mmap_threshold']} | {p['baseline_lane']} / {p['candidate_lane']} | {p['median_percent']['api_sum_ns']:+.3f}% | {p['counter_percent']['instructions']:+.3f}% | {p['counter_percent']['cycles']:+.3f}% | {a['counters']['page-faults']:,} / {b['counters']['page-faults']:,} |")
    lines+=['','| Ordinary process | Build | API p50 ms | Minor faults |','|---|---|---:|---:|']
    for r in d['ordinary_process_faults']:lines.append(f"| {r['path']} | {r['build']} | {r['api_p50_ns']/1e6:.3f} | {r['minor_faults']:,} |")
    return '\n'.join(lines)+'\n'
if __name__=='__main__':
    d=derive();(ROOT/'diagnostic-summary.json').write_text(json.dumps(d,indent=2)+'\n');(ROOT/'diagnostic-summary.md').write_text(render(d));print(json.dumps({'status':'pass','processes':8,'samples':240}))
