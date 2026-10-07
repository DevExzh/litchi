#!/usr/bin/env python3
import hashlib,json,datetime
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0943')
def pairs(items):
    result={}
    for k,v in items:
        assert k not in result, f'duplicate key: {k}'
        result[k]=v
    return result
def read(p): return json.loads(p.read_text(),object_pairs_hook=pairs)
def desc(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for chunk in iter(lambda:f.read(1048576),b''):h.update(chunk)
    return dict(path=str(p.resolve()),bytes=p.stat().st_size,sha256=h.hexdigest())
records=[]; envs=[]
for version,source_name,receipt_name,build_suffix,target_name in [
    ('baseline','baseline-source-v1','baseline-source-v2.json','v1','target-baseline-v1'),
    ('candidate','candidate-source-v4','candidate-source-v4.json','v4','target-candidate-v4')]:
    source=ROOT/source_name; receipt_path=ROOT/receipt_name; receipt=read(receipt_path)
    assert receipt['root']==str(source)
    assert receipt['head']=='13562481317d3d219c6fcd93babd9377e84ea834'
    listing_path=Path(receipt['listing']['path']);assert desc(listing_path)==receipt['listing']
    rows=[json.loads(line,object_pairs_hook=pairs) for line in listing_path.read_text().splitlines()]
    assert len(rows)==receipt['files']==10798
    assert sum(row['bytes'] for row in rows)==receipt['bytes']
    assert {str(p.relative_to(source)) for p in source.rglob('*') if p.is_file()}=={x['path'] for x in rows}
    for row in rows:
        file=source/row['path'];assert desc(file)==dict(row,path=str(file)),str(file)
    if version=='candidate':
        assert receipt['changed_files']==['crates/litchi-doc/src/tracked_revision/package.rs']
        assert desc(Path(receipt['patch']['path']))==receipt['patch']
    for lane,bin_name in [('native','doc_retention_probe'),('allocator','doc_retention_probe_alloc')]:
        directory=ROOT/'builds'/f'{version}-{lane}-{build_suffix}'
        start=read(directory/'start.json');end=read(directory/'terminal.json')
        binary_receipt=read(directory/'binary.json'); binary=ROOT/target_name/'release'/bin_name
        manifest=source/'docs/performance/results/change-0917/probe/Cargo.toml'
        command=['/home/zhuhe/.cargo/bin/cargo','+1.95.0','build','--offline','--locked','-j2','--release','--message-format=json-render-diagnostics','--bin',bin_name,'--manifest-path',str(manifest),'--target-dir',str(ROOT/target_name)]
        assert start['argv']==end['argv']==command
        for record in [start,end]:
            assert record['cwd']==record['source_root']==str(source)
            assert record['source_receipt']==desc(receipt_path)
            assert record['manifest']==desc(manifest)
            assert record['lock']==desc(manifest.parent/'Cargo.lock')
            assert record['target_dir']==str(ROOT/target_name)
        assert end['exit_code']==0 and end['terminal_written_before_binary_descriptor'] is True
        assert end['binary_name']==bin_name
        assert start['started_utc']==end['started_utc']<=end['ended_utc']
        assert start['environment']==end['environment']
        env=start['environment'];envs.append(env)
        assert set(env)<=set(['HOME','CARGO_HOME','RUSTUP_HOME','PATH','LC_ALL','LANG','TZ','CARGO_TERM_COLOR','CARGO_PROFILE_RELEASE_DEBUG','CARGO_INCREMENTAL'])
        assert env['CARGO_PROFILE_RELEASE_DEBUG']=='1' and env['CARGO_INCREMENTAL']=='0' and env['CARGO_TERM_COLOR']=='never'
        assert binary_receipt==desc(binary)
        assert end['stdout']==desc(directory/'build.stdout.log') and end['stderr']==desc(directory/'build.stderr.log')
        records.append(dict(version=version,lane=lane,source_receipt=desc(receipt_path),start=desc(directory/'start.json'),terminal=desc(directory/'terminal.json'),binary=desc(binary),status='passed'))
assert all(env==envs[0] for env in envs)
result=dict(schema='litchi.performance.0943.independent-build-audit.v1',status='passed',utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),auditor=desc(Path(__file__)),sources_and_all_files_verified=True,exact_build_commands_environments_manifests_locks_verified=True,builds=records,timing_claim=False)
out=ROOT/'receipts/build-audit-v1.json'
with out.open('x') as f:f.write(json.dumps(result,indent=2)+'\n')
print(json.dumps(desc(out)))
