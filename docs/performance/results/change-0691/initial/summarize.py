#!/usr/bin/env python3
"""Summarize retained raw native samples without pooling instrumented results."""
import hashlib
import json
import math
import random
import statistics
from pathlib import Path

P = Path(__file__).resolve().parent
rng = random.Random(691)
def quantile(values, fraction):
    return sorted(values)[max(0, math.ceil(len(values)*fraction)-1)]
def bootstrap_median(values):
    n = len(values)
    samples = [statistics.median(rng.choices(values, k=n)) for _ in range(2000)]
    return [quantile(samples, .025), quantile(samples, .975)]
def parse(path):
    lines = path.read_text().splitlines()
    header = next(x for x in lines if x.startswith('sample\t')).split('\t')
    rows = [dict(zip(header, map(int, x.split('\t')))) for x in lines if x.split('\t')[0].isdigit()]
    assert len(rows) == 100
    assert [x['sample'] for x in rows] == list(range(100))
    metadata = {x.split('\t')[0]:x.split('\t')[1:] for x in lines if not x.split('\t')[0].isdigit()}
    return header[1:], rows, metadata

summary = []
semantics = {}
for run in json.loads((P / 'native-runs.json').read_text()):
    assert run['exit_code'] == 0
    path = P / run['output']
    assert hashlib.sha256(path.read_bytes()).hexdigest() == run['output_sha256']
    phases, rows, metadata = parse(path)
    stable = {k:v for k,v in metadata.items() if 'sha256' in k or k in ['target','before_slides','after_slides','before_shapes','after_shapes']}
    if run['case'] in semantics:
        assert stable == semantics[run['case']], run['case']
    semantics[run['case']] = stable
    for phase in phases:
        values = [r[phase] for r in rows]
        summary.append(dict(case=run['case'], leg=run['leg'], phase=phase, samples=len(values),
                            p50_ns=statistics.median(values), mean_ns=statistics.mean(values),
                            p95_ns=quantile(values,.95), p99_ns=quantile(values,.99),
                            min_ns=min(values), max_ns=max(values),
                            median_bootstrap_95ci_ns=bootstrap_median(values)))
(P / 'native-summary.json').write_text(json.dumps(summary,indent=2)+'\n')
(P / 'semantic-bindings.json').write_text(json.dumps(semantics,indent=2)+'\n')
for case in semantics:
    rows=[r for r in summary if r['case']==case and r['phase']=='total_ns']
    print(case, 'total p50 ms:', [round(r['p50_ns']/1e6,3) for r in rows])
print('Bootstrap is within-process descriptive uncertainty; independent-leg drift remains visible.')
