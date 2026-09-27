"""Independent exact-owner census from blank-separated native sample records."""
import gzip,json,re,sys
from pathlib import Path
p=Path(__file__).resolve().parent
analysis=json.loads((p/'native-analysis.json').read_text());rows=[]
owner=re.compile(r'^\s+[0-9a-f]+ namespace_uri_probe::capture_region_0793(?:\+0x[0-9a-f]+)? \(')
for row in analysis['processes']:
    path=p/'perf'/f"{row['repeat']}-{row['leg']}.frames.gz"
    blocks=[x for x in gzip.open(path,'rt').read().split('\n\n') if x.strip()]
    selected=[];unresolved=0
    for block in blocks:
        lines=block.splitlines();assert 'cycles:u:' in lines[0]
        matches=[i for i,line in enumerate(lines) if owner.match(line)]
        assert len(matches)<=1
        if matches:
            descendants=lines[1:matches[0]];selected.append(descendants)
            unresolved+=any('[unknown]' in line for line in descendants)
    assert len(blocks)==row['total_samples']
    assert len(selected)==row['owner_qualified_samples']
    assert unresolved==row['unresolved_owner_samples']
    rows.append({'repeat':row['repeat'],'leg':row['leg'],'total':len(blocks),'owner':len(selected),'unresolved_owner':unresolved})
result={'rows':rows,'phase_fraction_claim':False}
out=p/'root-native-counts.json'
if '--check' in sys.argv:assert json.loads(out.read_text())==result
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print('Independent native sample census PASS')
