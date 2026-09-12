"""Run CFB library guards before freezing candidate source."""
import json
import subprocess
import time
from run import HERE, REPO, TARGET, now, sha, write


def snapshot():
    names=json.loads((HERE/'baseline/source-manifest.json').read_text())
    return {name:sha(REPO/name) for name in names}


if __name__=='__main__':
    assert not (HERE/'preflight.receipt.json').exists()
    (TARGET/'test-tmp').mkdir(exist_ok=True)
    before=snapshot();write(HERE/'preflight-source-manifest.json',before)
    command=['env','TMPDIR='+str(TARGET/'test-tmp'),'CARGO_BUILD_JOBS=2','CARGO_INCREMENTAL=0','cargo','test',
        '--locked','-p','litchi-cfb','--all-features','--lib','--target-dir',str(TARGET)]
    start,tick=now(),time.monotonic()
    with (HERE/'preflight.stdout').open('x') as out,(HERE/'preflight.stderr').open('x') as err:
        result=subprocess.run(command,cwd=REPO,stdout=out,stderr=err)
    same=snapshot()==before
    receipt=dict(command=command,start_utc=start,end_utc=now(),seconds=time.monotonic()-tick,
        exit_code=result.returncode,source_unchanged=same,
        source_manifest_sha256=sha(HERE/'preflight-source-manifest.json'),
        artifacts={name:sha(HERE/name) for name in ['preflight.stdout','preflight.stderr']})
    write(HERE/'preflight.receipt.json',receipt)
    assert result.returncode==0 and same,receipt
    print('CFB preflight passed against unchanged source',flush=True)
