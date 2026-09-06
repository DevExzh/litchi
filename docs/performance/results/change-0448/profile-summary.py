#!/usr/bin/env python3
"""Derive weighted whole-process and run-frame self samples without profiler execution."""
import argparse,collections,gzip,hashlib,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent
HEADER=re.compile(r'^\S.*?\s+\d+\s+[0-9.]+:\s+(\d+)\s+cycles:u:\s*$')
FRAME=re.compile(r'^\s*[0-9a-f]+\s+(.+?)\s+\([^\n]*\)\s*$')
def raw(path):
    return path.read_bytes() if path.exists() else gzip.decompress(Path(str(path)+'.gz').read_bytes())
def profile(role):
    directory=ROOT/'profiles'/str({'separate-sleeps':1,'minimum-service':3}[role])/'record'
    receipt=json.loads((directory/'receipt.json').read_text());identity=receipt['artifacts']['perf_script'];data=raw(ROOT/identity['path'])
    assert len(data)==identity['bytes'] and hashlib.sha256(data).hexdigest()==identity['sha256']
    total=collections.Counter();run=collections.Counter();inclusive=collections.Counter();blocks=0;run_blocks=0;missing_blocks=0;missing_period=0;warnings=[]
    def consume(period,frames):
        nonlocal blocks,run_blocks,missing_blocks,missing_period
        if period is None:return
        if not frames:
            missing_blocks+=1;missing_period+=period;frames=['[missing callchain]']
        blocks+=1;total[frames[0]]+=period
        if any('litchi_perf_baseline::pptx_provider_lifecycle::run_lifecycle_iteration' in frame for frame in frames):
            run_blocks+=1;run[frames[0]]+=period
            for frame in set(frames):
                if "litchi_opc" in frame or "soapberry_zip" in frame:inclusive[frame]+=period
    period=None;frames=[]
    for line in data.decode().splitlines():
        header=HEADER.match(line)
        if header:
            consume(period,frames);period=int(header.group(1));frames=[]
        elif not line.strip():
            consume(period,frames);period=None;frames=[]
        elif period is not None and (frame:=FRAME.match(line)):
            frames.append(re.sub(r'\+0x[0-9a-f]+$','',frame.group(1)))
        elif line.strip():warnings.append(line)
    consume(period,frames)
    assert blocks>0 and sum(total.values())>0
    def rows(counter, weight=None):
        weight=sum(counter.values()) if weight is None else weight
        return [{'symbol':symbol,'period':period,'percent':100*period/weight} for symbol,period in sorted(counter.items(),key=lambda item:(-item[1],item[0]))[:25]] if weight else []
    return {'role':role,'sample_blocks':blocks,'missing_callchain_blocks':missing_blocks,'missing_callchain_period':missing_period,'total_period':sum(total.values()),'whole_process_self':rows(total),'run_frame_blocks':run_blocks,'run_frame_period':sum(run.values()),'run_frame_percent':100*sum(run.values())/sum(total.values()),'run_frame_self':rows(run),'run_frame_inclusive':rows(inclusive,sum(run.values())),'unparsed_lines':warnings,'scope':'cycles:u weighted self; run-frame subset includes setup/probes/warmups/report work inside the run function, not just elapsed intervals; inclusive rows overlap and must not be added'}
def derive():return {'change':448,'profiles':[profile(mode) for mode in ('separate-sleeps','minimum-service')],'scope':'diagnostic CPU attribution for a pacing model; blocked sleep is absent from cycles; no exact timed-region or production CPU improvement claim'}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--check',action='store_true');a=p.parse_args();v=derive();path=ROOT/'profile-summary.json'
    if a.check:assert json.loads(path.read_text())==v
    else:
        with path.open('x') as stream:stream.write(json.dumps(v,indent=2)+'\n')
    print('VALID')
