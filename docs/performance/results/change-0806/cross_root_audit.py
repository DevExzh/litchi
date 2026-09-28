"""Independent 0806 numerical replay from raw cross-format sample arrays."""
import json, random, statistics, sys
from pathlib import Path
import cross_analysis as cross
p=Path(__file__).resolve().parent
read=lambda n:json.loads(Path(n).read_text())
builds=cross.verify_build_receipts()
source_chain=cross.verify_candidate_transition(builds)
rows={}
for receipt in read(p/'cross-native/receipts.json'):
    report=read(receipt['report']['path'])
    for result in report['results']:
        samples=sorted(result['elapsed_ns']['samples'])
        assert len(samples)==30
        midpoint=(samples[14]+samples[15])//2
        assert midpoint==result['elapsed_ns']['p50']
        rows[(result['case'],result['corpus']['shape'],receipt['block'],receipt['leg'])]=midpoint
output=[]
analysis=read(p/'cross-analysis.json')
assert analysis['candidate_transition']==source_chain
for row in analysis['rows']:
    case,shape=row['case'],row['shape']
    ratios=[rows[(case,shape,b,'after')]/rows[(case,shape,b,'before')] for b in range(6)]
    rng=random.Random(806081)
    estimates=sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)]) for _ in range(10000))
    ratio=statistics.median(ratios);lo,hi=estimates[250],estimates[9749]
    assert ratios==row['ratios'] and ratio==row['ratio_median']
    assert (lo,hi)==(row['bootstrap']['ci_low'],row['bootstrap']['ci_high'])
    assert row['reject']==(ratio>1.05 and lo>1)
    output.append({'case':case,'shape':shape,'ratio':ratio,'ci_low':lo,'ci_high':hi,'reject':row['reject']})
assert len(output)==8 and len(rows)==96
result={'raw_process_rows':len(rows),'rows':output,
        'rejected':any(r['reject'] for r in output),
        'candidate_transition':source_chain}
out=p/'cross-root-audit.json'
if '--check' in sys.argv:assert read(out)==result
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print('Independent cross-format numerical audit PASS')
