#!/usr/bin/env python3
"""Retain a diagnostic caller-boundary calibration before formal profiles."""
import importlib.util
import json
from pathlib import Path
import subprocess
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
spec=importlib.util.spec_from_file_location('B',P/'build.py');B=importlib.util.module_from_spec(spec);spec.loader.exec_module(B)
def main():
    d=P/'profile-calibration';d.mkdir()
    source=B.census();assert source==json.loads((P/'source-baseline.json').read_text())
    (d/'source-before.json').write_text(json.dumps(source,indent=2)+'\n')
    owner='litchi_docx::package::codec::<impl litchi_docx::package::model::Package>::write_plain'
    binary=B.BIN/'baseline-native'
    command=['taskset','-c','12','valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(d/'raw.callgrind'),str(binary),'--case','docx_ordinary_save_counting_publish','--samples','1','--warmup','0','--filesystem-root',str(ROOT.parent/'litchi-0709-fs'),'--json',str(d/'output.json')]
    with (d/'stdout').open('w') as out,(d/'stderr').open('w') as err:r=subprocess.run(command,cwd=ROOT,stdout=out,stderr=err)
    assert B.census()==source
    (d/'source-after.json').write_text(json.dumps(B.census(),indent=2)+'\n')
    record=dict(purpose='Callgrind raw caller calibration only; excluded from final repeated attribution',command=command,exit_code=r.returncode,binary_sha256=B.sha(binary),binary_bytes=binary.stat().st_size,build_sha256=B.sha(P/'build-baseline.json'),script_sha256=B.sha(Path(__file__)),source_manifest_sha256=B.sha(P/'source-baseline.json'),artifacts={f.name:B.sha(f) for f in d.iterdir()})
    (d/'receipt.json').write_text(json.dumps(record,indent=2)+'\n');print('Profile calibration',r.returncode,flush=True);assert r.returncode==0
if __name__=='__main__':main()
