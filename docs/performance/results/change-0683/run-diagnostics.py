#!/usr/bin/env python3
"""Serial native long-inline ABBA diagnostic, separate from allocator instrumentation."""
import hashlib
import json
from pathlib import Path
import subprocess

packet = Path(__file__).resolve().parent
out = packet / 'diagnostics'
out.mkdir(exist_ok=True)
records = []
for leg, suffix in [('a1','before'),('b1','after'),('b2','after'),('a2','before')]:
    binary = Path('/home/zhuhe/code/litchi-target-0683-' + suffix) / 'release/xlsx0683'
    corpus = Path('/home/zhuhe/code/litchi-0683-corpus/long-inline.xlsx')
    command = ['perf','stat','-x,','-e','cycles,instructions,cache-misses,page-faults',
               '-o',str(out / (leg + '.csv')),'taskset','-c','12',str(binary),
               'bench','cells-selected',str(corpus),'Sheet1','3','1000']
    with (out / (leg + '.tsv')).open('w') as stdout, (out / (leg + '.stderr')).open('w') as stderr:
        result = subprocess.run(command,stdout=stdout,stderr=stderr)
    records.append(dict(command=command,exit_code=result.returncode,
        binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest(),
        corpus_sha256=hashlib.sha256(corpus.read_bytes()).hexdigest()))
    (out / 'commands.json').write_text(json.dumps(records,indent=2)+'\n')
    if result.returncode: raise SystemExit(result.returncode)
raw = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in out.iterdir() if p.is_file()}
(out / 'raw-sha256.json').write_text(json.dumps(raw,indent=2)+'\n')
