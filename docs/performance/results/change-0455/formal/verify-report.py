#!/usr/bin/env python3
"""Validate retained sink observations, then replay the frozen lifecycle oracle."""
import copy,importlib.util,json,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('lifecycle',ROOT/'lifecycle-oracle.py');base=importlib.util.module_from_spec(spec);spec.loader.exec_module(base)
def require(ok,message):
    if not ok:raise ValueError(message)
def check_report(value):
    report=copy.deepcopy(value)
    for row in report['samples_raw']:
        sink=row.pop('publication_sink')
        require(set(sink)=={'accepted_bytes','write_calls','largest_write','write_size_buckets'},'sink fields')
        for field in ['accepted_bytes','write_calls','largest_write']:require(type(sink[field]) is int and 0<=sink[field]<2**64,'sink integer')
        require(sink['accepted_bytes']==row['output_bytes']==report['expected_output_bytes'],'sink accepted bytes')
        require(0<sink['largest_write']<=65536,'sink write cap')
        buckets=sink['write_size_buckets'];bounds={'bytes_0':(0,0),'bytes_1_to_512':(1,512),'bytes_513_to_4096':(513,4096),'bytes_4097_to_16384':(4097,16384),'bytes_16385_to_65536':(16385,65536),'bytes_over_65536':(65537,2**64-1)}
        require(set(buckets)==set(bounds),'sink histogram fields')
        require(all(type(n) is int and 0<=n<2**64 for n in buckets.values()),'sink histogram integer')
        require(sum(buckets.values())==sink['write_calls'],'sink histogram count')
        require(buckets['bytes_over_65536']==0,'sink oversized write')
        lower=sum(bounds[k][0]*n for k,n in buckets.items());upper=sum(bounds[k][1]*n for k,n in buckets.items())
        require(lower<=sink['accepted_bytes']<=upper,'sink histogram byte bounds')
    return base.check_report(report)
if __name__=='__main__':print(json.dumps(check_report(json.loads(Path(sys.argv[1]).read_text())),sort_keys=True))
