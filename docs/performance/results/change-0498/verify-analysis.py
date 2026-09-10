#!/usr/bin/env python3
"""Recompute every latency/throughput comparison independently of analysis code."""
import csv
import json
import math
import re
from pathlib import Path

HERE=Path(__file__).resolve().parent
analysis=json.loads((HERE/'final-analysis.json').read_text())
checked=0
for comparison in analysis['comparisons']:
    data={}
    for side in ['batch','serial']:
        path=HERE/'final'/comparison[side+'_csv']
        data[side]=[r for r in csv.DictReader(path.open()) if r['record']=='sample' and int(r['index'])>=3]
    scopes=[(None,comparison['comparison']['aggregate'])]+[(int(k),v) for k,v in comparison['comparison']['repeats'].items()]
    for repeat,values in scopes:
        expected_flags=[]
        for side in ['batch','serial']:
            rows=[r for r in data[side] if repeat is None or int(r['repeat'])==repeat]
            ns=sorted(int(r['elapsed_ns']) for r in rows)
            expected={'mean_us':sum(ns)/len(ns)/1000}
            expected.update({f'p{p}_us':ns[math.ceil(len(ns)*p/100)-1]/1000 for p in [50,95,99]})
            for metric,value in expected.items():
                assert math.isclose(values['latency'][metric][side],value,rel_tol=1e-12)
            throughput=sum(int(r['logical_bytes']) for r in rows)*1e9/sum(ns)
            assert math.isclose(values['throughput'][side+'_bytes_s'],throughput,rel_tol=1e-12)
        for metric,value in values['latency'].items():
            delta=(value['batch']/value['serial']-1)*100
            assert math.isclose(value['delta_pct'],delta,abs_tol=1e-9)
            if delta>5:expected_flags.append('latency_'+metric+'_regression_gt_5pct')
        t=values['throughput'];delta=(t['batch_bytes_s']/t['serial_bytes_s']-1)*100
        assert math.isclose(t['delta_pct'],delta,abs_tol=1e-9)
        if delta < -5: expected_flags.append('throughput_regression_gt_5pct')
        if repeat is None:
            rss=values['rss']
            for side in ['batch','serial']:
                stderr=(HERE/'final'/comparison[side+'_csv'].replace('.csv','.time.stderr')).read_text()
                expected=int(re.search(r'Maximum resident set size \(kbytes\):\s*(\d+)',stderr).group(1))
                assert rss[side+'_kib']==expected
            delta=(rss['batch_kib']/rss['serial_kib']-1)*100
            assert math.isclose(rss['delta_pct'],delta,abs_tol=1e-9)
            if delta>5:expected_flags.append('rss_regression_gt_5pct')
        assert sorted(expected_flags)==sorted(values['adverse_flags'])
        checked+=1
for item in analysis['scaling']:
    agg=item['scaling']['aggregate'] if 'scaling' in item else item['aggregate']
    if item['corpus']=='few-large' and item['workers']>4:assert not agg['amdahl_meaningful']
result={'status':'pass','comparison_scopes_recomputed':checked,'latency_throughput_rss_flags_verified':True}
(HERE/'analysis-verification.json').write_text(json.dumps(result,indent=2)+'\n')
print(json.dumps(result))
