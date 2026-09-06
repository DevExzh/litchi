#!/usr/bin/env python3
"""Replay deterministic I/O contract evidence; these are not latency samples."""
import argparse,gzip,hashlib,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def load(n):return json.loads((ROOT/n).read_text())
def raw(n):
    p=ROOT/n
    return p.read_bytes() if p.exists() else gzip.decompress(Path(str(p)+'.gz').read_bytes())
def require(ok,label):
    if not ok:raise ValueError(label)
def derive():
    receipt=load('checks/final-zip-tests.json');record=receipt['log'];data=raw(record['path'])
    require(receipt['status']=='pass' and receipt['exit_code']==0 and receipt['passed_tests']==455,'final ZIP tests')
    require(hashlib.sha256(data).hexdigest()==record['sha256'] and len(data)==record['bytes'],'test log custody')
    rows=[json.loads(m) for m in re.findall(rb'FUSED_CAPTURE_IO (\{[^\n]+\})',data)]
    require([(r['deflated'],r['decoded_bytes']) for r in rows]==[(d,s) for d in [False,True] for s in [0,1,65536,262181,1048613]],'exact fixture matrix')
    for r in rows:
        require(all(type(v) is int and 0<=v<2**64 for k,v in r.items() if k!='deflated'),'integer counters')
        require(r['fused_calls']<r['control_calls'] and r['control_bytes']>=r['fused_bytes']+r['compressed_bytes'],'source pass reduction')
        r['returned_bytes_saved']=r['control_bytes']-r['fused_bytes']
        r['returned_bytes_relative_percent']=100*(r['fused_bytes']/r['control_bytes']-1)
    return {'change':450,'decision':'retain measured ZIP enabler','rows':rows,'cases':10,'new_tests':5,'zip_tests':455,'independent_fresh_indexes':True,'source_max_chunk':65536,'generator':'xorshift32 seed 0x5a17932d; Store/Deflate single-member StreamingArchiveWriter; exact tested source retained','scope':'deterministic I/O assertions from one unit-test execution, not 10 timing samples; input index construction excluded, strict layout proof included independently; no latency, allocation, PPTX integration or native/scaling claim'}
def render(value):
    lines=['# Combined ZIP capture/decode I/O evidence','','Each alternative uses an independently indexed identical immutable archive.','Counters begin after indexing and include each path\'s strict layout validation.','These ten deterministic cases are correctness/I/O assertions, not latency samples.','','| Method | Decoded bytes | Compressed bytes | Control calls / bytes | Combined calls / bytes | Returned-byte change |','| --- | ---: | ---: | ---: | ---: | ---: |']
    for r in value['rows']:
        lines.append(f"| {'Deflate' if r['deflated'] else 'Store'} | {r['decoded_bytes']:,} | {r['compressed_bytes']:,} | {r['control_calls']:,} / {r['control_bytes']:,} | {r['fused_calls']:,} / {r['fused_bytes']:,} | {r['returned_bytes_relative_percent']:.3f}% |")
    lines+=['','Control reads decoded bytes and then captures/compares compressed bytes. Combined','captures once, decodes once and returns both validated outputs. Every case preserves','decoded bytes and exact compressed token data; the token survives dropping the','source archive and publishes through the preservation writer without recompression.','The pre-existing destination member survives that publication.','','Both operations retain complete decoded and compressed payloads. No allocation','count, peak-memory reduction, latency or production PPTX speedup is established.','Empty and one-byte cases include substantial metadata overhead; larger cases remove','roughly half the returned source bytes. OPC/PPTX adoption still requires explicit','budget, cache, token lifetime, semantic-validation and source-freshness integration.','']
    return '\n'.join(lines)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();v=derive();m=render(v)
    if a.check:require(load('measurements.json')==v and (ROOT/'measurements.md').read_text()==m,'derived measurements')
    else:
        with (ROOT/'measurements.json').open('x') as f:f.write(json.dumps(v,indent=2)+'\n')
        (ROOT/'measurements.md').write_text(m)
    print(json.dumps({'status':'pass','cases':10,'zip_tests':455,'check':a.check}))
