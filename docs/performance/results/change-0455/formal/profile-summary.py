#!/usr/bin/env python3
"""Summarize separate whole-process perf stat observations without API attribution."""
import csv,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
proof=json.loads((ROOT/'profile-proof.json').read_text());rows=[]
for i,record in enumerate(proof['rows']):
    counters={}
    for row in csv.reader((ROOT/'profiles'/str(i)/'perf.csv').read_text().splitlines()):
        if not row or row[0].startswith('#'):continue
        count=int(row[0]);event=row[2];counters[event]={'count':count,'event_runtime_ns':int(row[3]),'running_percent':float(row[4])}
    assert len(counters)==6
    rows.append({'lane':i,'build':record['build'],'counters':counters,'ipc':counters['instructions']['count']/counters['cycles']['count']})
(ROOT/'profile-summary.json').write_text(json.dumps({'scope':proof['scope'],'rows':rows},indent=2)+'\n');print(json.dumps({'status':'pass','processes':len(rows)}))
