#!/usr/bin/env python3
"""Derive deterministic OPC returned-byte assertions; these are not timings."""
import gzip,hashlib,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
def require(condition,message):
    if not condition:raise ValueError(message)
def raw(name):
    p=ROOT/name
    return p.read_bytes() if p.exists() else gzip.decompress(p.with_name(p.name+'.gz').read_bytes())
def load(name):return json.loads(raw(name))
def derive():
    log=raw('checks/final-opc-tests.log').decode()
    rows=[json.loads(x) for x in re.findall(r'OPC_COMBINED_IO (\{[^\n]+\})',log)]
    require(len(rows)==8,'eight I/O cases')
    require([(r['deflated'],r['decoded_bytes']) for r in rows]==[(f,n) for f in (False,True) for n in (0,1,65536,262181)],'complete ordered cases')
    for r in rows:
        require(all(type(r[k]) is int and r[k]>=0 for k in r if k!='deflated'),'nonnegative counters')
        require(r['combined_calls']<r['control_calls'],'fewer calls')
        require(r['control_bytes']-r['combined_bytes']>=r['compressed_bytes'],'one compressed pass removed')
        if not r['deflated']:require(r['compressed_bytes']==r['decoded_bytes'],'Store identity')
    return {'change':451,'kind':'deterministic positional I/O assertions','latency_samples':0,'rows':rows}
def render(value):
    lines=['# OPC combined capture: deterministic I/O assertions','','Identical, independently opened packages; counters begin after opening and end','after cold read/authorization. Maximum provider return is 65,536 bytes. Tests','also compare exact decoded/compressed bytes and whole published outputs.','','| Method | Decoded bytes | Compressed bytes | Control calls | Combined calls | Control returned bytes | Combined returned bytes |','|---|---:|---:|---:|---:|---:|---:|']
    for r in value['rows']:
        lines.append('| '+('Deflate' if r['deflated'] else 'Store')+' | '+' | '.join(str(r[k]) for k in ['decoded_bytes','compressed_bytes','control_calls','combined_calls','control_bytes','combined_bytes'])+' |')
    return '\n'.join(lines)+'\n\nNo latency, RSS, allocation-peak or complete PPTX speedup claim.\n'
if __name__=='__main__':
    v=derive();(ROOT/'measurements.json').write_text(json.dumps(v,indent=2)+'\n');(ROOT/'measurements.md').write_text(render(v))
