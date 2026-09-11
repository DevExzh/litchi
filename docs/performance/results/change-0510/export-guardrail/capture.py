#!/usr/bin/env python3
"""Serial matched public-export newline guardrail; separate counting-sink scope."""
import csv, datetime, json, subprocess, sys, time
from pathlib import Path
HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parent))
from run import REPO, SCRATCH, sha, sources, write
for stage,repeat in [('before','r1'),('after','r1'),('after','r2'),('before','r2')]:
    build=json.loads((HERE/f'{stage}-build.json').read_text())
    binary=SCRATCH/stage/'export-guardrail'
    assert sha(binary)==build['binary_sha256']
    assert sha(HERE/'probe.rs')==build['probe_source_sha256']
    current=sources()
    assert current==json.loads((HERE.parent/'source-manifest.json').read_text())
    command=['taskset','-c','2',str(binary),'500']
    started=datetime.datetime.now(datetime.timezone.utc).isoformat();tick=time.monotonic()
    with (HERE/f'{stage}-{repeat}.csv').open('x') as out,(HERE/f'{stage}-{repeat}.log').open('x') as err:
        result=subprocess.run(['/usr/bin/time','-v',*command],cwd=REPO,stdout=out,stderr=err)
    unchanged=sources()==current
    write(HERE/f'{stage}-{repeat}-receipt.json',{'command':command,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':result.returncode,'source_unchanged':unchanged,'binary_sha256':sha(binary),'build_receipt_sha256':sha(HERE/f'{stage}-build.json'),'library_source_manifest_sha256':build['library_source_manifest_sha256'],'samples_per_case':500,'warmups':10,'cpu_affinity':[2],'scope':'public export to zero-retaining byte-counting sink; opening and exact byte checks outside timer; no timed hashing','artifacts':{n:sha(HERE/n) for n in [f'{stage}-{repeat}.csv',f'{stage}-{repeat}.log']}})
    assert result.returncode==0 and unchanged and sha(binary)==build['binary_sha256']
    with (HERE/f'{stage}-{repeat}.csv').open() as stream:rows=list(csv.DictReader(stream))
    assert len(rows)==4500
    print(stage,repeat,'4500 samples',flush=True)
