#!/usr/bin/env python3
"""Capture A/A before edits, then A/B/B/A with frozen native binaries."""
import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
mode=sys.argv[1]
assert mode in ['baseline','compare']
sources={
 'real':str(ROOT/'test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx'),
 'control':str(P/'marker-control.pptx'),
 'generated':'generated:12x8',
 'notes-poi':str(ROOT/'test-data/poi/test-data/slideshow/prProps.pptx'),
 'notes-lo':str(ROOT/'test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx'),
}
cases=[(workflow+'-'+name,source,workflow) for workflow in ['one','noop','two']
       for name,source in sources.items() if workflow!='two' or name in ['real','control','generated']]
legs=[('a0','baseline'),('a1','baseline')] if mode=='baseline' else [
      ('a2','baseline'),('b0','candidate'),('b1','candidate'),('a3','baseline')]
records=[]
out=P/'native';out.mkdir(exist_ok=True)
for leg_index,(leg,phase) in enumerate(legs):
    binary=ROOT.parent/'litchi-0701-bin'/f'{phase}-native'
    ordered=cases if leg_index%2==0 else list(reversed(cases))
    for name,source,workflow in ordered:
        command=['taskset','-c','12',str(binary),'phases',source,'100','5',workflow]
        output=out/f'{name}-{leg}.tsv';error=output.with_suffix('.stderr')
        started=time.monotonic()
        with output.open('w') as stdout,error.open('w') as stderr:
            result=subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
        records.append(dict(case=name,workflow=workflow,leg=leg,phase=phase,command=command,
                            exit_code=result.returncode,seconds=time.monotonic()-started,
                            output=str(output.relative_to(P)),
                            output_sha256=hashlib.sha256(output.read_bytes()).hexdigest(),
                            stderr_sha256=hashlib.sha256(error.read_bytes()).hexdigest(),
                            binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),
                            source_sha256=None if source.startswith('generated:') else hashlib.sha256(Path(source).read_bytes()).hexdigest()))
        (P/f'native-runs-{mode}.json').write_text(json.dumps(records,indent=2)+'\n')
        print(name,leg,result.returncode,flush=True)
        assert result.returncode==0,error.read_text()
