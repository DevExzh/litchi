#!/usr/bin/env python3
"""Independent XML start/declaration census; not production execution counts."""
import json,hashlib
from pathlib import Path
from xml.parsers import expat
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
manifest=json.loads((P/'corpus/manifest.json').read_text())
rows=[]
for member in manifest['members']:
 path=ROOT/member['corpus_path'];counts=dict(starts=0,starts_with_namespace_declarations=0,namespace_declarations=0)
 def start(name,attrs):
  count=sum(k=='xmlns' or k.startswith('xmlns:') for k in attrs)
  counts['starts']+=1
  counts['starts_with_namespace_declarations']+=bool(count)
  counts['namespace_declarations']+=count
 parser=expat.ParserCreate();parser.StartElementHandler=start;parser.Parse(path.read_bytes(),True)
 rows.append(dict(uri=member['uri'],sha256=hashlib.sha256(path.read_bytes()).hexdigest(),**counts))
by_uri={r['uri']:r for r in rows}
weighted={k:sum(by_uri[call['uri']][k] for call in manifest['calls']) for k in ['starts','starts_with_namespace_declarations','namespace_declarations']}
(P/'element-counts.json').write_text(json.dumps(dict(method='Python Expat raw QName start events; source syntax, not MCE branch execution instrumentation',members=rows,weighted_capture_sequence=weighted),indent=2)+'\n')
print(weighted)
