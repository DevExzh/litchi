#!/usr/bin/env python3
"""Serial source-bound XLS corpus, allocation, logical-I/O, and native AA/ABBA capture."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import time

PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PHASE, BEFORE, AFTER, ALLOC_BEFORE, ALLOC_AFTER = sys.argv[1:]
assert PHASE in {'baseline', 'candidate'}
OUT = PACKET / 'measurements' / PHASE
OUT.mkdir(parents=True, exist_ok=True)
CPU = '12'
WARMUPS, SAMPLES = 3, 30
binary = BEFORE if PHASE == 'baseline' else AFTER
allocator = ALLOC_BEFORE if PHASE == 'baseline' else ALLOC_AFTER
source_root = Path('/home/zhuhe/code/litchi-0684-before') if PHASE == 'baseline' else ROOT
commands = []
def sha(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def run(command, name):
    command = ['taskset', '-c', CPU, *map(str, command)]
    start = time.monotonic()
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
    (OUT / name).write_text(result.stdout)
    if result.stderr: (OUT / (name + '.stderr')).write_text(result.stderr)
    commands.append(dict(command=command, output=name, exit_code=result.returncode, seconds=time.monotonic()-start))
    (OUT/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
    if result.returncode: raise SystemExit(f'{name}: {result.returncode}: {result.stderr}')
    return json.loads(result.stdout)

manifest = dict(phase=PHASE, head=subprocess.check_output(['git','rev-parse','HEAD'],cwd=source_root,text=True).strip(), cpu=CPU, warmups=WARMUPS, samples=SAMPLES,
    binary_sha256=sha(binary), allocator_binary_sha256=sha(allocator),
    source_sha256={str(p.relative_to(source_root)):sha(p) for p in sorted((source_root/'crates/litchi-xls').rglob('*.rs'))},
    probe_sha256={str(p.relative_to(PACKET)):sha(p) for folder in ['probe','allocation-probe'] for p in sorted((PACKET/folder).rglob('*')) if p.is_file() and '__pycache__' not in str(p)},
    corpus_manifest_sha256=sha(PACKET/'corpus-manifest.json'))
(OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
for mode in ['owned','file']:
    corpus=run([binary,'corpus','--root','test-data','--mode',mode,'--sample-coordinates','16','--max-queries','512'],f'corpus-{mode}.json')
    assert corpus['query_mismatches']==0,(mode,corpus['query_mismatches'])
    if PHASE=='candidate': assert corpus==json.loads((PACKET/'measurements/baseline'/f'corpus-{mode}.json').read_text()),mode
    print(PHASE,'corpus',mode,corpus['files_seen'],flush=True)

if PHASE=='baseline':
    corpus=json.loads((OUT/'corpus-owned.json').read_text())
    cases=[]
    targets=['test-data/poi/test-data/spreadsheet/54016.xls','test-data/poi/test-data/spreadsheet/45365-2.xls','test-data/ole/xls/WithCustomViews.xls','test-data/poi/test-data/spreadsheet/15228.xls','test-data/poi/test-data/spreadsheet/Simple.xls']
    for path in targets:
        file=next(f for f in corpus['files'] if 'test-data/'+f['path']==path)
        sheet=next(s for s in file['worksheets'] if s['status']=='ok' and s['distinct_coordinates']>0)
        first=next(q for q in sheet['queries'] if 'first-stored' in q['reasons'])
        last=next(q for q in sheet['queries'] if 'last-stored' in q['reasons'])
        cases.append(dict(case=Path(path).stem+'-stored',path=path,sha256=file['sha256'],sheet=sheet['worksheet'],row=first['row'],column=first['column'],second_row=last['row'],second_column=last['column']))
        if Path(path).stem in ['54016','WithCustomViews','Simple']:
            missing=next(q for q in sheet['queries'] if q['expected'] is None and q['error'] is None)
            cases.append(dict(case=Path(path).stem+'-missing',path=path,sha256=file['sha256'],sheet=sheet['worksheet'],row=missing['row'],column=missing['column'],second_row=first['row'],second_column=first['column']))
    file=next(f for f in corpus['files'] if f['path']=='ole/xls/ConditionalFormattingSamples.xls')
    cases.append(dict(case='formula-refusal',path='test-data/'+file['path'],sha256=file['sha256'],sheet=6,row=0,column=0,second_row=0,second_column=1))
    (PACKET/'cases.json').write_text(json.dumps(cases,indent=2)+'\n')
else: cases=json.loads((PACKET/'cases.json').read_text())
manifest['cases_sha256']=sha(PACKET/'cases.json')
def route_args(case, route, mode, samples=SAMPLES, warmups=WARMUPS):
    return ['route','--input',case['path'],'--route',route,'--mode',mode,'--worksheet',str(case['sheet']),'--row',str(case['row']),'--column',str(case['column']),'--second-row',str(case['second_row']),'--second-column',str(case['second_column']),'--warmups',str(warmups),'--samples',str(samples)]
for case in cases:
    for mode in ['owned','file']:
        run([binary,*route_args(case,'prepared',mode,1,1)],f'counts-{case["case"]}-{mode}.json')
        if case['case']!='formula-refusal':
            for op in ['q1','q2','q3','visit']:
                for repeat in range(3):
                    run([allocator,mode,op,case['path'],case['sheet'],case['row'],case['column']],f'alloc-{case["case"]}-{mode}-{op}-{repeat}.json')
    print(PHASE,'counts/alloc',case['case'],flush=True)
for case in cases:
    for mode in ['owned-native','file-native']:
        for route in ['prepared','visit']:
            legs=[('a1',BEFORE),('a2',BEFORE)] if PHASE=='baseline' else [('a1',BEFORE),('b1',AFTER),('b2',AFTER),('a2',BEFORE)]
            for leg,program in legs:
                run([program,*route_args(case,route,mode)],f'{leg}-{case["case"]}-{mode}-{route}.json')
    print(PHASE,'timing',case['case'],flush=True)
manifest['raw_sha256']={p.name:sha(p) for p in sorted(OUT.iterdir()) if p.is_file() and p.name!='manifest.json'}
(OUT/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(PHASE,'complete',flush=True)
