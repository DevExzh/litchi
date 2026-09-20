#!/usr/bin/env python3
"""Verify producer totals and whole-child RSS without mixing their scopes."""
import json
import math
import re
import statistics
from capture import HERE, sha, source_check

def read(name):
    receipt = json.loads((HERE / (name+'.receipt.json')).read_text())
    build = json.loads((HERE / 'build.json').read_text())
    assert receipt['exit_code'] == 0
    assert receipt['binary_sha256'] == build['binary_sha256']
    assert receipt['script_sha256'] == sha(HERE / 'capture.py')
    for path, digest in receipt['artifacts'].items():
        assert sha(HERE / path) == digest
    raw = json.loads((HERE / (name+'.json')).read_text())
    assert raw['binary_identity']['binary_sha256'] == build['binary_sha256']
    assert raw['tool']['instrumentation'] == 'none'
    assert len(raw['results']) == 1
    return raw['results'][0]

def main():
    source_check()
    producer = []
    first = None
    for repeat in [1, 2]:
        name = f'producer-r{repeat}'
        r = read(name)
        assert r['case'] == 'xlsx_producer_medium_source_one_edit_save'
        e = r['elapsed_ns']; v = e['samples']; n = len(v)
        assert n == 100 and sorted(v) == v and sorted(e['sample_order']) == list(range(n))
        computed = dict(p50=int(statistics.median(v)),p95=v[math.ceil(n*.95)-1],
                        p99=v[math.ceil(n*.99)-1],mean=statistics.mean(v),min=min(v),max=max(v))
        for key, value in computed.items():
            assert math.isclose(e[key], value, rel_tol=1e-12)
        identity = {key:r[key] for key in ['corpus','sink','output_sha256']}
        if first is None: first = identity
        else: assert identity == first
        producer.append(dict(name=name, elapsed_ns=computed))
    rss = []
    for shape in ['medium', 'dense-sparse']:
        for repeat in [1, 2]:
            name = f'rss-r{repeat}-{shape}'
            r = read(name)
            native = json.loads((HERE / f'native-r1-{shape}.json').read_text())['results'][0]
            for key in ['corpus','sink','output_sha256']:
                assert r[key] == native[key]
            text = (HERE / (name+'.rss.txt')).read_text()
            assert re.search(r'Exit status:\s+0\b', text)
            peak = int(re.search(r'Maximum resident set size \(kbytes\):\s+(\d+)', text)[1])
            rss.append(dict(name=name, whole_child_peak_rss_kib=peak))
    result = dict(producer=producer,producer_identity=first,rss=rss,
                  scope='Producer combined native interval; RSS is whole child including setup and oracles. No historical speed comparison.')
    (HERE/'control-analysis.json').write_text(json.dumps(result,indent=2)+'\n')
    print('PASS producer repetition and whole-child RSS controls')

if __name__ == '__main__': main()
