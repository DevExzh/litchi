#!/usr/bin/env python3
"""Replay attribution and complete imported custody without original workspace."""
import argparse, hashlib, importlib.util, json
from pathlib import Path
ROOT=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('derive',ROOT/'derive.py');d=importlib.util.module_from_spec(spec);spec.loader.exec_module(d)
def check(sealed):
    ledger=d.load('inputs.json');origin=(ROOT/'inputs/origin-SHA256SUMS.txt').read_bytes()
    d.require(d.digest(origin)==ledger['origin_seal_sha256'],'origin seal identity')
    seal={n:h for h,n in (line.split('  ',1) for line in origin.decode().splitlines())}
    build=d.load('inputs/build.json');manifest=d.load('inputs/'+build['source_manifest']['path'])
    raw=(ROOT/'inputs'/build['source_manifest']['path']).read_bytes()
    d.require(d.digest(raw)==build['source_manifest']['sha256'] and len(manifest)==build['source_manifest']['files'],'build manifest')
    d.require(len(ledger['records'])==35 and len({r['path'] for r in ledger['records']})==35,'import count/uniqueness')
    for r in ledger['records']:
        if r['path'].startswith('inputs/'):
            name=r['path'][7:];d.require(r['sha256']==seal[name],'sealed origin member')
            d.require(r['origin']=='docs/performance/results/change-0448/'+name,'origin path')
        else:
            d.require(r['path']=='source/'+r['origin'] and manifest[r['origin']]==r['sha256'],'build source identity')
    d.require(d.derive()==d.load('attribution.json'),'recomputed attribution')
    spec=importlib.util.spec_from_file_location('render',ROOT/'render.py');renderer=importlib.util.module_from_spec(spec);spec.loader.exec_module(renderer)
    d.require((ROOT/'measurements.md').read_text()==renderer.render(),'rendered measurements')
    decision=d.load('decision.json');d.require(decision['decision']=='retain diagnostic attribution' and decision['production_optimization'] is False and decision['goal_complete'] is False and decision['new_workload_executions']==0,'scoped decision')
    probes=d.load('probes.json');d.require(probes['status']=='pass' and probes['parser_and_classification_checks']==7 and len(probes['rejected'])==6 and probes['temporary_directory_removed'],'probe receipt')
    if sealed:
        records={}
        for line in (ROOT/'SHA256SUMS').read_text().splitlines():
            h,n=line.split('  ',1);p=Path(n)
            d.require(not p.is_absolute() and '..' not in p.parts and n not in records,'safe seal name')
            d.require((ROOT/p).resolve().is_relative_to(ROOT.resolve()),'seal path escape')
            d.require(d.digest((ROOT/p).read_bytes())==h,'sealed member '+n);records[n]=h
        d.require(set(records)=={str(p.relative_to(ROOT)) for p in ROOT.rglob('*') if p.is_file() and p.name!='SHA256SUMS'},'complete seal')
    return {'change':449,'status':'pass','reports_reanalysed':8,'samples_reanalysed':240,'profiles_reanalysed':2,'sealed':sealed,'new_workload_executions':0}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--sealed',action='store_true');a=p.parse_args()
    try: print(json.dumps(check(a.sealed)))
    except Exception as error: print('INVALID: '+str(error));raise SystemExit(1)
