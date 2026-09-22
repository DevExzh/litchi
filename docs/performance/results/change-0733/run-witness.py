#!/usr/bin/env python3
"""Fixed standalone Callgrind collection-state witness; no library benchmark."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
BIN=ROOT.parent/'litchi-0733-bin'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
out=P/'witness';out.mkdir();BIN.mkdir()
source=P/'collection-witness.rs';binary=BIN/'collection-witness'
commands=[['rustfmt','--check',str(source)],['rustc','--edition=2024','--crate-name','collection_witness','-C','opt-level=2','-C','debuginfo=1','-D','warnings',str(source),'-o',str(binary)],['taskset','-c','12',str(binary)]]
for repeat in range(3):
 commands.append(['taskset','-c','12','valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect=collection_witness::owner','--callgrind-out-file='+str(out/f'{repeat}.callgrind'),'--log-file='+str(out/f'{repeat}.valgrind'),str(binary)])
write(P/'witness-plan.json',dict(scope='collection counter semantics only; no PPT timing evidence',source_sha256=sha(source),runner_sha256=sha(Path(__file__)),commands=commands))
rows=[]
for index,command in enumerate(commands):
 r=subprocess.run(command,cwd=ROOT,capture_output=True);stdout=out/f'{index}.stdout';stderr=out/f'{index}.stderr';stdout.write_bytes(r.stdout);stderr.write_bytes(r.stderr);rows.append(dict(command=command,exit_code=r.returncode,stdout=stdout.name,stdout_sha256=sha(stdout),stderr=stderr.name,stderr_sha256=sha(stderr)));write(out/'manifest.json',dict(status='running',runs=rows));assert r.returncode==0,r.stderr.decode()
write(out/'manifest.json',dict(status='passed',runs=rows,source_sha256=sha(source),binary=dict(path=str(binary),bytes=binary.stat().st_size,sha256=sha(binary)),profiles={str(f.relative_to(P)):sha(f) for f in out.iterdir() if f.suffix in ['.callgrind','.valgrind']}))
print('PASS standalone native control and three collection-window witnesses')
