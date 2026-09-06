#!/usr/bin/env python3
"""Require exact baseline/candidate pilot artifact and scalar identities."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
rows = []
for shape in ('tiny', 'medium', 'large'):
    before = json.loads((ROOT / 'before/pilots' / (shape+'.json')).read_text())['results'][0]
    after = json.loads((ROOT / 'after/pilots' / (shape+'.json')).read_text())['results'][0]
    for field in ('corpus', 'output_sha256', 'sink', 'source'):
        assert before[field] == after[field], (shape, field)
    rows.append({'shape': shape, 'status': 'pass', 'output_sha256': before['output_sha256'],
                 'target_payload_sha256': before['corpus']['target_payload_sha256'],
                 'semantic_sha256': before['source']['ods_scalar_rows']['semantic_sha256']})
print(json.dumps({'status':'pass','scope':'pilot exact artifact, sink and source summary identity; no formal timing claim','rows':rows},indent=2))
