#!/usr/bin/env python3
"""Derive same-owner combined hashing fractions for the written report."""
import hashlib
import json
from pathlib import Path
import statistics

P = Path(__file__).resolve().parent

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def read(p):
    return json.loads(p.read_text())

manifest = read(P / 'captures/manifest.json')
rows = []
for run in manifest['runs']:
    if run['route'] != 'profiled-clock':
        continue
    path = P / 'captures' / run['output']
    assert sha(path) == run['sha256']
    samples = read(path)['samples']
    values = []
    for sample in samples:
        spans = sample['diagnostics']['commit']['spans']
        duration = sum(s['duration_ns'] for s in spans if s['phase'] in
                       ['ArtifactHashBefore', 'ArtifactHashAfter'])
        values.append((duration, duration / sample['whole_ns'],
                       duration / sample['split']['commit_ns']))
    rows.append(dict(cycle=run['cycle'], repeat=run['repeat'],
                     combined_hash_p50_ns=statistics.median(v[0] for v in values),
                     combined_hash_whole_ratio_p50=statistics.median(v[1] for v in values),
                     combined_hash_commit_ratio_p50=statistics.median(v[2] for v in values)))
result = dict(scope='observed route only; combine spans per owner before taking medians',
              capture_manifest_sha256=sha(P / 'captures/manifest.json'),
              script_sha256=sha(Path(__file__)), processes=rows)
(P / 'report-stats.json').write_text(json.dumps(result, indent=2) + '\n')
for key in rows[0]:
    if key not in ['cycle', 'repeat']:
        print(key, min(r[key] for r in rows), max(r[key] for r in rows))
