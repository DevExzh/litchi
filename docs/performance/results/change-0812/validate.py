"""Replay fresh instruction localization, independent counts, and rebuild refusal."""
import ast
import gzip
import hashlib
from pathlib import Path
import subprocess
import sys
import inputs as c

def artifact(row):
    assert c.artifact(row['path']) == row, row['path']
    return Path(row['path'])

def interval(row):
    assert row['exit_code'] == 0 and row['started'] <= row['ended']
    return row['started'], row['ended']

for script in ('instruction_analysis.py', 'offset_audit.py', 'fresh_offset_audit.py'):
    subprocess.run([sys.executable, '-B', str(c.P / script), '--check'], cwd=c.ROOT, check=True)
custody=c.verify()
a=c.read(c.P/'instruction-analysis.json');b=c.read(c.P/'fresh-offset-audit.json')
assert a['binary']==b['binary'] and a['measured_outputs_verified']==200
assert not a['historical_offsets_mapped']
assert len(a['rows'])==len(b['rows'])==2
for left,right in zip(a['rows'],b['rows'],strict=True):
    assert left['repeat']==right['repeat']
    assert left['whole_samples']==right['whole_process_samples']
    assert left['owner_samples']==right['owner_qualified_samples']
    assert left['leaf_samples']==right['scanner_leaf_samples']
    assert left['leaf_period']==right['scanner_leaf_period']
    assert left['unknown_interior']==right['unknown_interior_samples']
    assert left['lost_event_lines']==right['lost_event_lines']
    assert {r['offset_hex']:(r['samples'],r['period']) for r in left['offsets']}=={
        r['offset_hex']:(r['samples'],r['period']) for r in right['offsets']}
r=c.read(c.P/'rebuild/receipt.json');f=c.read(c.P/'rebuild/prefreeze.json')
commands=c.read(c.P/'rebuild/commands.json')
assert commands['build']==r['build']
assert r['status']=='binary-mismatch' and not r['historical_mapping_authorized']
assert r['assembly'] is None and commands['nm'] is None and commands['objdump'] is None
assert r['binary']==a['binary'] and r['binary_comparison']['bytes_equal']
assert not r['binary_comparison']['sha256_equal'] and not r['binary_comparison']['exact']
old=c.read(c.OLD/'build/build.json')['binaries']['fp']
assert r['binary_comparison']['expected']==old and r['binary_comparison']['actual']==a['binary']
assert old['sha256']!=a['binary']['sha256'] and old['bytes']==a['binary']['bytes']
assert r['source']['hashes_equal'] and r['source']['file_count']==9196
assert r['build']['source_files_equal'] and r['build']['exit_code']==0
assert c.read(r['source']['before']['path'])['files']==c.read(c.OLD/'build/source.json')['files']
assert r['build']['command']==f['commands']['build']==['cargo','build','--offline','--locked','--release',
    '--manifest-path',str(c.OLD/'probe-src/Cargo.toml'),'--features','capture-profile']
assert r['build']['environment']==f['environment']
e=f['environment']
assert e['CARGO_TARGET_DIR']==str(c.ROOT.parent/'litchi-target-0811')
assert e['CARGO_BUILD_JOBS']=='2' and e['CARGO_INCREMENTAL']=='0'
assert e['RUSTFLAGS']=='-C force-frame-pointers=yes'
assert all(not v for v in e['inherited_compiler_environment'].values())
for row in commands['tool_versions'].values():
    assert interval(row)[1]<=r['build']['started']
assert r['started']<=r['build']['started']<=r['build']['ended']<=r['ended']
fresh=c.P/'fresh';ff=c.read(fresh/'frozen.json');complete=c.read(fresh/'complete.json')
assert ff['inputs']==complete['inputs']==custody and ff['binary']==complete['binary']==a['binary']
artifact(ff['build'])
for row in ff['drivers'].values():artifact(row)
for row in ff['tool_versions'].values():assert row['exit_code']==0
captures=c.read(fresh/'receipts.json');decodes=c.read(fresh/'decode.json')
assert len(captures)==len(decodes)==2 and complete['reports']==2 and complete['samples']==200
last=r['ended']
for i,row in enumerate(captures):
    assert row==c.read(fresh/f'{i}.receipt.json') and row['repeat']==i
    assert row['binary']==a['binary']
    assert interval(row)[0]>=last;last=row['ended']
    assert row['command']==['taskset','-c','12','perf','record','--no-buildid-cache','-e','cycles:u','-F','499',
        '--call-graph','fp','-o',str(fresh/f'{i}.data'),'--',a['binary']['path'],'--mode','capture','--shape','large',
        '--samples','100','--warmup','0','--output',str(fresh/f'{i}.json')]
    artifact(row['report']);artifact(row['output'])
for i,row in enumerate(decodes):
    assert row==c.read(fresh/f'{i}.decode.json') and row['repeat']==i
    assert row['binary']==a['binary'] and row['raw']==captures[i]['raw']
    assert interval(row)[0]>=last;last=row['ended']
    assert row['command']==['perf','script','--no-inline','--ns','-i',str(fresh/f'{i}.data')]
    artifact(row['errors'])
compression=c.read(fresh/'compression.json')
assert {(v['repeat'],v['kind']) for v in compression}=={(i,k) for i in range(2) for k in ('data','frames')}
assert len(compression)==4
for row in compression:
    stored=artifact(row['compressed']).read_bytes();raw=gzip.decompress(stored)
    assert stored[4:8]==b'\0\0\0\0'
    assert hashlib.sha256(raw).hexdigest()==row['original']['sha256'] and len(raw)==row['original']['bytes']
    assert not Path(row['original']['path']).exists()
    i=row['repeat'];expected=captures[i]['raw'] if row['kind']=='data' else decodes[i]['output']
    assert row['original']==expected
resume=c.read(fresh/'assembly-resume.json');artifact(resume['driver']);artifact(resume['prior_nm_stdout'])
assert resume['workload_retries']==0 and resume['started']>=last
for name in ('nm','objdump'):
    row=c.read(fresh/f'{name}.json');assert interval(row)[0]>=last;last=row['ended']
    artifact(row['output']);assert not row['stderr'] and row['binary']==a['binary']
assert (fresh/'symbols.txt').read_bytes()==(fresh/'symbols-resumed.txt').read_bytes()
assert c.read(fresh/'nm.json')['command']==['nm','-S','--defined-only',a['binary']['path']]
base,size=complete['address'],complete['size']
assert c.read(fresh/'objdump.json')['command']==['objdump','-d','--demangle','--no-show-raw-insn',
    f'--start-address=0x{base:x}',f'--stop-address=0x{base+size:x}',a['binary']['path']]
assert complete['ended']>=last
line=c.read(fresh/'line-mapping.json')
assert line['binary']==a['binary'] and not line['stderr'] and interval(line)[0]>=complete['ended']
assert line['command']==['addr2line','-afiC','-e',a['binary']['path'],hex(base+0x24d),hex(base+0x27d)]
assert 'core::result::Result<T,E>::map_err' in line['stdout'] and '/notes/codec.rs:389' in line['stdout']
cleanup=c.P/'cleanup.json'
if '--final' in sys.argv or cleanup.exists():
    removed=c.read(cleanup)
    assert removed['target_removed'] and not Path(removed['target']).exists()
    assert removed['target']==e['CARGO_TARGET_DIR'] and removed['binary']==a['binary']
    assert line['ended']<=removed['started']<=removed['ended'] and removed['inputs']==custody
else:
    artifact(a['binary'])
assert not list(c.P.rglob('__pycache__'))
for path in c.P.glob('*.py'):ast.parse(path.read_text())
print('0812 aggregate PASS: historical mapping refused; fresh two-report/200-output localization and independent counts agree')
