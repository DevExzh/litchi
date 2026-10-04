#!/usr/bin/env python3
"""Root-owned serial qualification/build runner; all phase receipts create-only."""
from pathlib import Path
import datetime, hashlib, json, os, shutil, subprocess, sys, tarfile
P = Path(__file__).resolve().parent
R = P.parents[3]
OLD = P.parent / 'change-0894'
S = R / 'target/litchi-0898-source'
T = R / 'target/litchi-0898-build'
CARGO = '/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo'
MARKER = '.owner-litchi-0898'
PATHS = ('crates/litchi-ooxml-common/src/mce/codec.rs', 'crates/litchi-ooxml-common/src/mce/tests.rs')
SUPPORT = ('tools/check_crate_boundaries.py','tools/crate_boundaries.json','docs/adr/0008-migration-and-verification.md','docs/adr/0028-iwa-monolith-exit.md')

def read(p): return json.loads(p.read_text())
def sha(p):
    h=hashlib.sha256()
    with p.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest()
def desc(p):
    assert p.is_file() and not p.is_symlink(),p
    st=p.stat();return {'path':str(p.resolve()),'bytes':st.st_size,'sha256':sha(p),'mode':st.st_mode&0o777,'executable':bool(st.st_mode&0o100)}
def write(p,v):
    with p.open('x') as f: json.dump(v,f,indent=2,sort_keys=True);f.write('\n')
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
def env(release=False):
    e={'HOME':'/home/zhuhe','CARGO_HOME':'/home/zhuhe/.cargo','RUSTUP_HOME':'/home/zhuhe/.rustup','PATH':str(Path(CARGO).parent)+':/usr/bin:/bin','RUSTUP_TOOLCHAIN':'1.95.0','CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','CARGO_NET_OFFLINE':'true','CARGO_TARGET_DIR':str(T),'TMPDIR':str(T/'tmp'),'LANG':'C','LC_ALL':'C','TZ':'UTC','RUSTFLAGS':'','RUSTDOCFLAGS':'-D warnings','CARGO_PROFILE_DEV_DEBUG':'0','CARGO_PROFILE_TEST_DEBUG':'0'}
    if release:e.update(CARGO_PROFILE_RELEASE_OPT_LEVEL='3',CARGO_PROFILE_RELEASE_DEBUG='1',CARGO_PROFILE_RELEASE_LTO='fat',CARGO_PROFILE_RELEASE_CODEGEN_UNITS='1')
    return e

def run(label,argv,release=False):
    start=P/(label+'-start.json');log=P/(label+'.log');receipt=P/(label+'-run.json')
    assert not receipt.exists() and not log.exists()
    command={'argv':argv,'cwd':str(S),'environment':env(release),'started_utc':now(),'helper':desc(Path(__file__))}
    write(start,command)
    with log.open('x') as f:r=subprocess.run(argv,cwd=S,env=env(release),stdout=f,stderr=subprocess.STDOUT)
    v={**command,'finished_utc':now(),'exit_code':r.returncode,'status':'pass' if r.returncode==0 else 'fail','log':desc(log)}
    write(receipt,v);print(label,v['status'],flush=True);assert r.returncode==0,log
    return desc(receipt)

def prepare():
    assert not S.exists() and not T.exists()
    manifest=read(OLD/'source.json');archive=OLD/'source.tar.gz'
    assert desc(archive)['sha256']==manifest['archive']['sha256']
    selection=read(OLD/'source-start.json').get('selection')
    # The complete retained build source selection is unchanged at current HEAD.
    paths=['.cargo','Cargo.toml','README.md','clippy.toml','deny.toml','rust-toolchain.toml','rustfmt.toml','crates','test-data','tools/perf-baseline']
    proof=subprocess.check_output(['git','diff','506aa2edae','8d517a7861','--',*paths],cwd=R);assert not proof
    S.mkdir();T.mkdir();(T/'tmp').mkdir()
    for d in (S,T):(d/MARKER).write_text('0898 owned source/build\n')
    with tarfile.open(archive) as tar:
        for m in tar:
            assert m.isfile() or m.isdir();assert not Path(m.name).is_absolute() and '..' not in Path(m.name).parts
        tar.extractall(S,filter='data')
    for name,d in manifest['files'].items():assert sha(S/name)==d['sha256'],name
    for name in SUPPORT:
        dst=S/name;dst.parent.mkdir(parents=True,exist_ok=True);dst.write_bytes(subprocess.check_output(['git','show','8d517a7861:'+name],cwd=R))
    shutil.copy2(OLD/'quality-Cargo.lock',S/'Cargo.lock')
    shutil.copy2(S/'Cargo.lock',P/'quality-Cargo.lock')
    write(P/'source.json',{'status':'pass','baseline_revision':'506aa2edae6666d9c76ae6827c2f5941a8e47896','current_revision':'8d517a7861','source_selection_unchanged':True,'archive':desc(archive),'manifest':desc(OLD/'source.json'),'source_root':str(S),'target_root':str(T),'verified_files':len(manifest['files']),'root_lock':desc(P/'quality-Cargo.lock'),'standalone_lock':desc(S/'tools/perf-baseline/Cargo.lock'),'support':{n:desc(S/n) for n in SUPPORT}})
    print('prepare pass',flush=True)

def diagnostic(label):
    manifest=str(S/'tools/perf-baseline/Cargo.toml')
    run(label+'-diagnostic-build',[CARGO,'build','--manifest-path',manifest,'--locked','--offline','--bin','xml_attribute_bounds'])
    binary=P/(label+'-diagnostic');shutil.copy2(T/'debug/xml_attribute_bounds',binary)
    write(P/(label+'-diagnostic-identity.json'),desc(binary))
    roots=[S/'test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx',S/'test-data/ooxml/docx/drawing.docx',S/'test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx']
    run(label+'-differential',[str(binary),'differential','--json',str(P/(label+'-differential.json')),'--generated','20000','--aliasing','20000',*[str(f) for f in roots]])

def baseline():
    previous=P.parent/'change-0896'
    identity=read(previous/'baseline-diagnostic-identity.json')
    original=previous/'baseline-diagnostic'
    assert sha(original)==identity['sha256']
    binary=P/'baseline-diagnostic';shutil.copy2(original,binary)
    write(P/'baseline-diagnostic-identity.json',desc(binary))
    prior=read(previous/'baseline-diagnostic-build.json');assert prior['status']=='pass' and prior['exit_code']==0
    write(P/'baseline-diagnostic-build-run.json',{'status':'pass','exit_code':0,'kind':'reuse_exact_qualified_baseline_binary','binary':desc(binary),'original_binary':desc(original),'original_build':desc(previous/'baseline-diagnostic-build.json'),'source':desc(P/'source.json'),'reason':'Baseline selected source unchanged from506aa2; reuse avoids an identical rebuild. Fresh differential executes both roles against the same0898source fixtures.'})
    roots=[S/'test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx',S/'test-data/ooxml/docx/drawing.docx',S/'test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx']
    run('baseline-differential',[str(binary),'differential','--json',str(P/'baseline-differential.json'),'--generated','20000','--aliasing','20000',*[str(f) for f in roots]])

def freeze():
    assert not (P/'candidate-source.json').exists()
    overlay=P/'candidate-overlay';overlay.mkdir()
    source={}
    for n in PATHS:
        f=R/n;dst=overlay/n;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(f,dst);shutil.copy2(f,S/n);source[n]=desc(dst)
    patch=subprocess.check_output(['git','diff','--binary','8d517a7861','--',*PATHS],cwd=R);assert patch
    (P/'candidate.patch').write_bytes(patch)
    write(P/'candidate-source.json',{'status':'pass','source':desc(P/'source.json'),'overlay':source,'patch':desc(P/'candidate.patch'),'production_adopted':False,'guard_contract':'post-endpoint-per-worksheet-tolerance-v1'})
    print('candidate frozen',flush=True)

def check_source():
    v=read(P/'candidate-source.json')
    for n,d in v['overlay'].items():assert sha(S/n)==d['sha256']==sha(R/n),n
    manifest=read(OLD/'source.json')
    for n,d in manifest['files'].items():
        if n not in v['overlay']:assert sha(S/n)==d['sha256'],n
    assert sha(S/'Cargo.lock')==sha(P/'quality-Cargo.lock')

def quality():
    check_source();m=str(S/'crates/litchi-ooxml-common/Cargo.toml');base=['--manifest-path',m,'--locked','--offline'];perf=['--manifest-path',str(S/'tools/perf-baseline/Cargo.toml'),'--locked','--offline']
    gates=[('common-tests',[CARGO,'test',*base,'--all-features']),('common-clippy',[CARGO,'clippy',*base,'--all-features','--all-targets','--','-D','warnings']),('common-doc',[CARGO,'doc',*base,'--all-features','--no-deps']),('xlsx-tests',[CARGO,'test','--manifest-path',str(S/'crates/litchi-xlsx/Cargo.toml'),'--locked','--offline','--lib']),('docx-tests',[CARGO,'test','--manifest-path',str(S/'crates/litchi-docx/Cargo.toml'),'--locked','--offline','--lib']),('pptx-tests',[CARGO,'test','--manifest-path',str(S/'crates/litchi-pptx/Cargo.toml'),'--locked','--offline','--lib']),('diagnostic-tests',[CARGO,'test',*perf,'--bin','xml_attribute_bounds']),('producer-tests',[CARGO,'test',*perf,'--lib','producer_shape::']),('fmt',[str(Path(CARGO).parent/'rustfmt'),'--check','--edition','2024',*[str(S/n) for n in PATHS]]),('boundary',['python3','-B',str(S/'tools/check_crate_boundaries.py')])]
    receipts={}
    for name,argv in gates:receipts[name]=run(name+"-v2",argv)
    check_source();write(P/'quality.json',{'status':'pass','source':desc(P/'candidate-source.json'),'gates':receipts,'gate_count':len(receipts)})

def release():
    check_source();bins=P/'binaries';bins.mkdir(exist_ok=True);roles={}
    for role,binary,features in [('native','litchi-perf-baseline',[]),('allocator','litchi-perf-baseline-alloc',['allocator-metrics'])]:
        argv=[CARGO,'build','--manifest-path',str(S/'tools/perf-baseline/Cargo.toml'),'--release','--locked','--offline','--bin',binary,'-j','2']
        if features:argv+=['--features',','.join(features)]
        receipt=run('candidate-'+role+'-build',argv,True);dst=bins/('candidate-'+role);shutil.copy2(T/'release'/binary,dst)
        dep=P/('candidate-'+role+'.d');shutil.copy2(T/'release'/(binary+'.d'),dep)
        elf=subprocess.check_output(['/usr/bin/readelf','-n',str(dst)],text=True);buildid=next(l.split('Build ID:')[1].strip() for l in elf.splitlines() if 'Build ID:' in l)
        roles[role]={**desc(dst),'name':binary,'role':role,'features':features,'elf_build_id':buildid,'build_receipt':receipt,'depfile':desc(dep)}
    check_source();write(P/'candidate-build.json',{'status':'pass','source':desc(P/'candidate-source.json'),'quality':desc(P/'quality.json'),'binaries':roles,'environment':env(True),'compiler':subprocess.check_output([str(Path(CARGO).parent/'rustc'),'--version','--verbose'],text=True),'guard_contract':'post-endpoint-per-worksheet-tolerance-v1'})

if __name__=='__main__':
    action=sys.argv[1]
    if action=='prepare':prepare()
    elif action=='baseline':baseline()
    elif action=='freeze':freeze()
    elif action=='quality':quality()
    elif action=='candidate':diagnostic('candidate')
    elif action=='release':release()
    else:raise ValueError(action)
