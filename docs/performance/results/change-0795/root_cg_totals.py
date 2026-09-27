"""Independent scalar raw-cost conservation, without the function-graph parser."""
import re,json,sys
from pathlib import Path
p=Path(__file__).resolve().parent
rows=[]
for raw in sorted((p/'profiles').glob('*.callgrind*')):
    events=None;positions=None;summary=None;totals=None;pending_call=False;self_sum=None;cost_rows=0;call_rows=0
    for line in raw.read_text().splitlines():
        if line.startswith('events:'):
            events=line.split()[1:];self_sum=[0]*len(events)
        elif line.startswith('positions:'):positions=line.split()[1:]
        elif line.startswith('summary:'):summary=[int(x) for x in line.split()[1:]]
        elif line.startswith('totals:'):totals=[int(x) for x in line.split()[1:]]
        elif line.startswith('calls='):
            assert not pending_call;pending_call=True
        elif re.match(r'^[*+\-0-9]',line):
            assert events and positions
            fields=line.split();values=[int(x) for x in fields[len(positions):]]
            assert len(values)<=len(events) and all(x>=0 for x in values)
            values += [0]*(len(events)-len(values))
            if pending_call:call_rows+=1;pending_call=False
            else:self_sum=[a+b for a,b in zip(self_sum,values)];cost_rows+=1
    assert not pending_call and events==['Ir','Bc','Bcm','Bi','Bim']
    assert summary is not None and totals is not None
    summary += [0]*(len(events)-len(summary));totals += [0]*(len(events)-len(totals))
    assert self_sum==summary==totals,(raw.name,self_sum,summary,totals)
    assert (summary[0]>0)==raw.name.endswith('.1')
    rows.append({'file':raw.name,'totals':dict(zip(events,summary)),'self_rows':cost_rows,'call_rows':call_rows})
assert len(rows)==24
result={'rows':rows,'scope':'Independent raw counter conservation only; no latency/cycle attribution'}
out=p/'root-cg-totals.json'
if '--check' in sys.argv:assert json.loads(out.read_text())==result
else:assert not out.exists();out.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print('Independent raw Callgrind counter conservation PASS')
