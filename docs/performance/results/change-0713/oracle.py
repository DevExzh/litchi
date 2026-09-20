#!/usr/bin/env python3
"""Run the identical public scanner oracle against both frozen source states."""
import importlib.util,json,shutil,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent;O=P/'oracle'
spec=importlib.util.spec_from_file_location('c0713',P/'custody.py');C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def main():
    stage=sys.argv[1];assert stage in ['baseline','candidate'];out=O/stage;out.mkdir(exist_ok=True);assert not (out/'result.json').exists()
    source=C.census();assert source==json.loads((P/f'source-{stage}.json').read_text())
    probe={n:C.sha(O/n) for n in ['Cargo.toml','Cargo.lock','src/main.rs']}
    (out/'source.json').write_text(json.dumps(source,indent=2)+'\n');(out/'probe.json').write_text(json.dumps(probe,indent=2)+'\n')
    cmd=['cargo','build','--release','--locked','--manifest-path',str(O/'Cargo.toml'),'--target-dir',str(C.TARGET),'-j','2'];start=time.monotonic()
    with (out/'build.log').open('w') as f:r=subprocess.run(cmd,cwd=C.ROOT,stdout=f,stderr=subprocess.STDOUT)
    record=dict(command=cmd,exit_code=r.returncode,seconds=time.monotonic()-start,source_sha256=C.sha(out/'source.json'),probe_sha256=C.sha(out/'probe.json'),log_sha256=C.sha(out/'build.log'))
    (out/'build.json').write_text(json.dumps(record,indent=2)+'\n');assert r.returncode==0
    binary=C.BIN/(stage+'-oracle');shutil.copy2(C.TARGET/'release/docx-active-offset-oracle-0712',binary)
    cmd=[str(binary),'--output',str(out/'report.json')]
    with (out/'stdout').open('w') as stdout,(out/'stderr').open('w') as stderr:r=subprocess.run(cmd,cwd=C.ROOT,stdout=stdout,stderr=stderr)
    assert C.census()==source and {n:C.sha(O/n) for n in probe}==probe
    result=dict(command=cmd,exit_code=r.returncode,binary=dict(path=str(binary),sha256=C.sha(binary),bytes=binary.stat().st_size),source_sha256=C.sha(out/'source.json'),probe_sha256=C.sha(out/'probe.json'),artifacts={n:C.sha(out/n) for n in ['report.json','stdout','stderr'] if (out/n).exists()})
    (out/'result.json').write_text(json.dumps(result,indent=2)+'\n');assert r.returncode==0
    report=json.loads((out/'report.json').read_text());assert report['case_count']==len(report['cases']) and len({x['name'] for x in report['cases']})==report['case_count']
    if stage=='candidate':
        assert probe==json.loads((O/'baseline/probe.json').read_text())
        assert (out/'report.json').read_bytes()==(O/'baseline/report.json').read_bytes(),'oracle outcomes differ'
    print(stage,'oracle PASS',report['case_count'],'cases',flush=True)
if __name__=='__main__':main()
