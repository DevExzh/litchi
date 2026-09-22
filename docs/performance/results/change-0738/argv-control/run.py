#!/usr/bin/env python3
"""Separate same-binary startup-argument controls; keep the main matrix intact."""
import copy
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time

S = Path(__file__).resolve().parent
sys.path.insert(0,str(S.parent))
from contract import P, ROOT, command, read, sha, validate_report, write
import run as parent_run
import analyze as primary
import audit as independent


def cmd(row):
    result = command(row)
    if row['arm']=='archive-extra':
        result += ['--operation','format']
    elif row['arm']=='prior-default':
        assert result[-2:]==['--lifecycle','legacy']
        result = result[:-2]
    return result


def guard():
    parent_run.guard()
    for path,digest in read(S/'freeze.json').items():
        assert sha(Path(path))==digest,path


def freeze():
    assert not (S/'freeze.json').exists()
    files = [S/'run.py',S/'plan.json',S/'hypothesis.md',P/'freeze.json']
    write(S/'freeze.json',{str(f):sha(f) for f in files})
    guard()


def analyze():
    guard()
    plan = read(S/'plan.json')
    manifest = read(S/'captures/manifest.json')
    assert manifest['status']=='complete'
    assert manifest['freeze_sha256']==sha(S/'freeze.json')
    assert manifest['preflight_sha256']==sha(S/'preflight.json')
    assert read(S/'preflight.json')['status']=='passed'
    assert read(S/'preflight.json')['freeze_sha256']==sha(S/'freeze.json')
    assert len(manifest['runs'])==len(plan['schedule'])==96
    processes=[]
    index={}
    expected_files={'manifest.json'}
    previous_end=0
    for actual,row in zip(manifest['runs'],plan['schedule']):
        assert {k:actual[k] for k in row}==row
        assert actual['command']==cmd(row) and actual['exit_code']==0
        assert actual['start_monotonic_ns']>=previous_end
        previous_end=actual['end_monotonic_ns']
        assert previous_end>actual['start_monotonic_ns']
        file=S/'captures'/actual['output'];stderr=file.with_suffix('.stderr')
        expected_files.update([file.name,stderr.name])
        assert sha(file)==actual['sha256'] and sha(stderr)==actual['stderr_sha256']
        assert stderr.read_bytes()==b''
        report=validate_report(read(file),row)
        item=dict(row)
        if row['lane']=='native':
            values=[s['phase_ns']['whole_ns'] for s in report['samples']]
            item['times_ns']=values
            item['stats']=primary.stats(values)
            independent.close(item['stats'],independent.statistics(values))
        else:
            item['allocation']=report['samples'][0]['allocations']['whole']
        key=(row['lane'],row['case'],row['arm'],row['repeat'])
        assert key not in index
        index[key]=item;processes.append(item)
    assert {f.name for f in (S/'captures').iterdir()}==expected_files
    comparisons=[];allocation=[]
    for case in ('primary','secondary'):
        for left,right in plan['comparisons']:
            pairs=[(index['native',case,left,i],index['native',case,right,i]) for i in range(9)]
            metrics={}
            for metric in ('p50','mean','p95','p99','maximum'):
                values=[primary.pct(a['stats'][metric],b['stats'][metric]) for a,b in pairs]
                metrics[metric]=primary.summarize(values)
                independently=[independent.percent(a['stats'][metric],b['stats'][metric]) for a,b in pairs]
                independent.close(metrics[metric],independent.summarize(independently,plan))
            comparisons.append(dict(case=case,left=left,right=right,metrics=metrics))
            pairs=[(index['allocation',case,left,i]['allocation'],index['allocation',case,right,i]['allocation']) for i in range(3)]
            allocation.append(dict(case=case,left=left,right=right,fields={f:dict(
                before=[a[f] for a,b in pairs],after=[b[f] for a,b in pairs],
                differences=[b[f]-a[f] for a,b in pairs]) for f in pairs[0][0]}))
    result=dict(status='passed',processes=processes,comparisons=comparisons,
                allocation_comparisons=allocation,native_samples=3600,
                arithmetic_check='Frozen independent scalar statistics/bootstrap helpers; grouping is shared here.',
                scope='Separate same-binary argument controls; no pooling with main matrix.')
    write(S/'analysis.json',result)
    print('PASS separate96-process argv matrix, exact oracles and independent scalar arithmetic')


def preflight():
    global S
    real=S
    guard()
    with tempfile.TemporaryDirectory(prefix='litchi-0738-argv-preflight-') as directory:
        S=Path(directory)
        for name in ('plan.json','freeze.json'):
            shutil.copy2(real/name,S/name)
        write(S/'preflight.json',dict(status='passed',synthetic=True,freeze_sha256=sha(S/'freeze.json')))
        (S/'captures').mkdir()
        rows=[]
        for n,row in enumerate(read(S/'plan.json')['schedule']):
            report=read(P/f'qualification-{row["build"]}'/f'{row["lane"]}-{row["case"]}-legacy.json')
            sample=report['samples'][0];report['samples_requested']=row['samples'];report['samples']=[]
            if row['build']!='archive':report['retained_witness_count']=row['samples']
            for i in range(row['samples']):
                item=copy.deepcopy(sample);item['index']=i
                if row['build']!='archive':item['retained_witness_count']=i+1
                if row['lane']=='native':item['phase_ns']['whole_ns']=1_000_000+n*1000+i*100
                report['samples'].append(item)
            name=f'{n:03d}.json';file=S/'captures'/name
            file.write_text(json.dumps(report,separators=(',',':'))+'\n');file.with_suffix('.stderr').write_bytes(b'')
            rows.append(dict(**row,command=cmd(row),exit_code=0,start_monotonic_ns=n*2+1,
                end_monotonic_ns=n*2+2,output=name,sha256=sha(file),stderr_sha256=sha(file.with_suffix('.stderr'))))
        write(S/'captures/manifest.json',dict(status='complete',freeze_sha256=sha(S/'freeze.json'),
            preflight_sha256=sha(S/'preflight.json'),runs=rows))
        analyze()
        manifest=read(S/'captures/manifest.json')
        manifest['runs'][0]['command'].append('--unexpected')
        write(S/'captures/manifest.json',manifest)
        try:
            analyze()
        except AssertionError:
            pass
        else:
            raise AssertionError('changed command accepted')
    S=real
    write(S/'preflight.json',dict(status='passed',synthetic=True,freeze_sha256=sha(S/'freeze.json'),
        altered_command_rejected=True,temporary_removed=True))


def capture():
    guard()
    pre=read(S/'preflight.json')
    assert pre['status']=='passed' and pre['freeze_sha256']==sha(S/'freeze.json')
    (S/'captures').mkdir()
    rows=[]
    for n,row in enumerate(read(S/'plan.json')['schedule']):
        file=S/'captures'/f'{n:03d}.json';stderr=file.with_suffix('.stderr');start=time.monotonic_ns()
        with file.open('wb') as out,stderr.open('wb') as err:
            result=subprocess.run(cmd(row),cwd=ROOT,stdout=out,stderr=err)
        end=time.monotonic_ns()
        rows.append(dict(**row,command=cmd(row),exit_code=result.returncode,start_monotonic_ns=start,
            end_monotonic_ns=end,output=file.name,sha256=sha(file),stderr_sha256=sha(stderr)))
        write(S/'captures/manifest.json',dict(status='running',freeze_sha256=sha(S/'freeze.json'),
            preflight_sha256=sha(S/'preflight.json'),runs=rows))
        assert result.returncode==0
        validate_report(read(file),row)
        print('PASS argv capture',n,row['case'],row['arm'],flush=True)
    guard()
    write(S/'captures/manifest.json',dict(status='complete',freeze_sha256=sha(S/'freeze.json'),
        preflight_sha256=sha(S/'preflight.json'),runs=rows))


if __name__=='__main__':
    {'freeze':freeze,'preflight':preflight,'capture':capture,'analyze':analyze}[sys.argv[1]]()
