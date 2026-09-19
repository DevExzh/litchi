#!/usr/bin/env python3
"""Serial process probe: baseline A/A, or candidate allocations and paired ABBA."""
import hashlib
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parents[5]
PACKET = Path(__file__).resolve().parent.parent
PHASE, BEFORE, AFTER, CORPUS = sys.argv[1:]
CORPUS = Path(CORPUS)
OUT = PACKET/'measurements'/PHASE
OUT.mkdir(parents=True, exist_ok=True)
CPU = '12'
SAMPLES = 20
WARMUP = 3
OPERATIONS = ['visit-cold', 'cells-cold', 'visit-selected', 'cells-selected', 'visit-warm', 'cells-warm']
cases = [(item['case'], CORPUS/(item['case']+'.xlsx'), 'Sheet1') for item in json.loads((CORPUS/'manifest.json').read_text())]
cases.append(('real-fallback', ROOT/'test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx', 'Лист 1'))
commands = []

def digest(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def run(command, name):
    start = time.monotonic()
    result = subprocess.run(['taskset', '-c', CPU, *map(str, command)], capture_output=True, text=True)
    (OUT/name).write_text(result.stdout)
    if result.stderr: (OUT/(name+'.stderr')).write_text(result.stderr)
    commands.append(dict(command=['taskset','-c',CPU,*map(str,command)], output=name, exit_code=result.returncode, seconds=time.monotonic()-start))
    (OUT/'commands.json').write_text(json.dumps(commands, indent=2)+'\n')
    if result.returncode: raise SystemExit(f'{name}: exit {result.returncode}: {result.stderr}')
    return result.stdout

binary = BEFORE if PHASE == 'baseline' else AFTER
alloc = Path(binary).with_name('xlsx0683_alloc')
manifest = dict(phase=PHASE, head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(), cpu=CPU, samples=SAMPLES, warmup=WARMUP, host=platform.platform(), rustc=subprocess.check_output(['rustc','-Vv'],text=True), binary_sha256=digest(binary), allocator_binary_sha256=digest(alloc), corpus={case:dict(path=str(path),sha256=digest(path),sheet=sheet) for case,path,sheet in cases}, source_sha256={str(p.relative_to(ROOT)):digest(p) for p in sorted((ROOT/'crates/litchi-xlsx').rglob('*.rs'))}, probe_sha256={str(p.relative_to(PACKET)):digest(p) for p in sorted((PACKET/'probe').rglob('*')) if p.is_file() and '__pycache__' not in str(p)})
# Baseline sources live in the isolated checkout, never the candidate worktree.
if PHASE == 'baseline':
    baseline_root=Path('/home/zhuhe/code/litchi-0683-before')
    manifest['source_sha256']={str(p.relative_to(baseline_root)):digest(p) for p in sorted((baseline_root/'crates/litchi-xlsx').rglob('*.rs'))}
    manifest['head']=subprocess.check_output(['git','rev-parse','HEAD'],cwd=baseline_root,text=True).strip()
(OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
run([binary,'layout'],'layout.tsv')
for case,path,sheet in cases:
    diff=run([binary,'diff',path,sheet],f'diff-{case}.tsv')
    if PHASE != 'baseline':
        assert diff == (PACKET/'measurements/baseline'/f'diff-{case}.tsv').read_text(), case
    if case != 'real-fallback':
        eligibility=run([binary,'scan',CORPUS/(case+'.xml')],f'scan-{case}.txt')
        expected='refused:' if case == 'malformed-escape' else ('not-eligible:' if case.startswith('late-') else 'eligible:')
        assert eligibility.startswith(expected), (case,eligibility)
    for operation in OPERATIONS:
        if case in ['late-refusal','malformed-escape'] and operation.endswith('-warm'): continue
        for repeat in range(3):
            run([alloc,operation,path,sheet],f'alloc-{case}-{operation}-{repeat}.tsv')
    print(PHASE,'matrix',case,flush=True)
for case,path,sheet in cases:
    for operation in OPERATIONS:
        if case in ['late-refusal','malformed-escape'] and operation.endswith('-warm'): continue
        legs=[('a1',BEFORE),('a2',BEFORE)] if PHASE == 'baseline' else [('a1',BEFORE),('b1',AFTER),('b2',AFTER),('a2',BEFORE)]
        for leg,executable in legs:
            run([executable,'bench',operation,path,sheet,WARMUP,SAMPLES],f'{leg}-{case}-{operation}.tsv')
    print(PHASE,'timing',case,flush=True)
# Bind every raw output after capture. No binaries need to be retained.
manifest['raw_sha256']={p.name:digest(p) for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='manifest.json'}
(OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(PHASE,'complete',flush=True)
