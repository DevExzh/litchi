#!/usr/bin/env python3
"""Derive confirmation statistics and selected whole-process profile counters."""
import argparse
import csv
import importlib.util
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def imported(name, path):
    spec = importlib.util.spec_from_file_location(name,path)
    value = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(value)
    return value


def derive():
    helper = imported('derive440', ROOT / 'derive.py')
    rows = []
    for phase in ('C1','C2','C3','C4'):
        directory = ROOT / 'confirmation' / (phase + ('-recapture' if phase=='C1' else ''))
        receipt = json.loads((directory/'receipt.json').read_text())
        assert receipt['status']=='pass'
        report = json.loads(helper.artifact(ROOT,receipt['artifacts']['report'],'confirmation'))
        elapsed = report['results'][0]['elapsed_ns']
        resource = helper.artifact(ROOT,receipt['artifacts']['resource_log'],'resource').decode()
        rss = int(re.search(r'Maximum resident set size \(kbytes\): (\d+)',resource)[1])*1024
        rows.append({'phase':phase,'role':receipt['role'],'elapsed_ns':helper.stats(elapsed['samples'],'elapsed',bootstrap=True),'rss_bytes':rss,'sample_order':elapsed['sample_order'],'elapsed_samples_ns':elapsed['samples']})
    pairs=[]
    for left,right in ((rows[0],rows[1]),(rows[3],rows[2])):
        for metric in ('p50','p95','p99','mean'):
            before,after=left['elapsed_ns'][metric],right['elapsed_ns'][metric]
            percent=helper.pct(before,after)
            pairs.append({'before_phase':left['phase'],'after_phase':right['phase'],'metric':metric,'before':before,'after':after,'relative_percent':percent,'flagged':percent>5})
        percent=helper.pct(left['rss_bytes'],right['rss_bytes'])
        pairs.append({'before_phase':left['phase'],'after_phase':right['phase'],'metric':'rss_bytes','before':left['rss_bytes'],'after':right['rss_bytes'],'relative_percent':percent,'flagged':percent>5})
    profiles={}
    for role in ('before','after'):
        stat=json.loads((ROOT/f'profiles/{role}/stat/formal/receipt.json').read_text())
        raw=helper.artifact(ROOT,stat['artifacts']['perf_stat'],'stat').decode()
        counters={}
        for fields in csv.reader(raw.splitlines()):
            if len(fields)>2 and fields[2] in stat['stat_events']:
                assert fields[2] not in counters
                counters[fields[2]]=int(fields[0])
        assert set(counters)==set(stat['stat_events'])
        path=ROOT/f'profiles/{role}/record/{"recapture" if role=="before" else "formal"}/receipt.json'
        record=json.loads(path.read_text())
        script=helper.artifact(ROOT,record['artifacts']['perf_script'],'script').decode()
        report=helper.artifact(ROOT,record['artifacts']['perf_report'],'perf report').decode()
        profiles[role]={'counters':counters,'ipc':counters['instructions:u']/counters['cycles:u'],'perf_script_addr2line_warnings':script.count('could not read first record'),'perf_report_addr2line_warnings':report.count('could not read first record'),'scope':'whole fresh process, including fixture setup, warmups and report gates; L1 zero has no cache-miss interpretation'}
    changes={name:helper.pct(profiles['before']['counters'][name],profiles['after']['counters'][name]) for name in profiles['before']['counters']}
    return {'change':440,'confirmation_rows':rows,'confirmation_comparisons':pairs,'confirmation_flags':[row for row in pairs if row['flagged']],'profiles':profiles,'profile_counter_change_percent':changes,'scope':'One fixed confirmation plan; original main tail flags and excluded overlap attempts remain retained. No automatic latency claim.'}


if __name__ == '__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args()
    value=derive();path=ROOT/'supplement.json'
    if args.check:
        assert json.loads(path.read_text())==value
    else:
        path.write_text(json.dumps(value,indent=2)+'\n')
    print('VALID: confirmation and profiles')
