#!/usr/bin/env python3
"""Switch only the owned content-type candidate, after serialized checks terminate."""
import argparse,datetime,hashlib,json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
p=argparse.ArgumentParser();p.add_argument('--role',choices=['before','after'],required=True);p.add_argument('--tag',required=True);args=p.parse_args()
for receipt in (ROOT/'checks').glob('*.json'):
    assert json.loads(receipt.read_text()).get('status')!='running',receipt
source=REPO/'crates/litchi-opc/src/content_type.rs'
variants={role:(ROOT/'candidate'/f'{role}-content_type.rs.txt').read_bytes() for role in ('before','after')}
previous=source.read_bytes();assert previous in variants.values(),'owned file differs from known source'
source.write_bytes(variants[args.role]);assert source.read_bytes()==variants[args.role]
row={'change':446,'status':'pass','role':args.role,'finished_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'path':str(source.relative_to(REPO)),'before_sha256':hashlib.sha256(previous).hexdigest(),'after_sha256':hashlib.sha256(source.read_bytes()).hexdigest()}
with (ROOT/'checks'/f'switch-{args.tag}.json').open('x') as stream:stream.write(json.dumps(row,indent=2)+'\n')
print(json.dumps(row))
