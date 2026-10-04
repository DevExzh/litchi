"""Commit the completed 0913 batch through a private index."""
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tempfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'
BRANCH='refs/heads/perf/0913-opc-owned-direct-read'
REPORT=ROOT/'docs/performance/0913-pptx-owned-direct-read.md'
INDEX_NAMES=('BASELINE.md','HOTSPOTS.md','REPORT.md','GOAL_AUDIT.md')
SOURCE_NAMES=('crates/soapberry-zip/src/office.rs','crates/litchi-opc/src/payload.rs','crates/soapberry-zip/tests/owned_direct_read.rs','crates/litchi-opc/tests/owned_direct_concurrency.rs','crates/soapberry-zip/tests/owned_direct_adversarial.rs')
def git(*args,env=None,data=None): return subprocess.check_output(['git',*args],cwd=ROOT,env=env,input=data)
def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def desc(path): return {'path':str(path.relative_to(ROOT)),'bytes':path.stat().st_size,'sha256':sha(path)}
def load(name): return json.loads((P/name).read_text())
def local_only(path):
    relative=path.relative_to(P)
    return relative.parts[0]=='binaries' or (relative.parts[0]=='capture' and path.suffix=='.pptx')
def main():
    assert not (P/'closure.json').exists()
    assert subprocess.run(['git','show-ref','--verify','--quiet',BRANCH],cwd=ROOT).returncode==1
    for name in ('decision.json','quality-summary.json','replay-before-cleanup.json','replay-after-cleanup.json','cleanup.json','result-review.json','independent-stats.json'):
        assert load(name)['status']=='pass',name
    assert (P/'replay-before-cleanup.json').read_bytes()==(P/'replay-after-cleanup.json').read_bytes()
    decision=load('decision.json'); adopted=decision['adopted']; assert type(adopted) is bool
    assert not adopted or load('analysis.json')['regression_flags']['owned_total_gate_pass']
    spec=importlib.util.spec_from_file_location('replay0913',P/'replay.py'); replay=importlib.util.module_from_spec(spec);spec.loader.exec_module(replay)
    review=load('result-review.json'); replay.check_descriptor_tree(review,'final review')
    assert review['final_report']==desc(REPORT)
    assert not list(P.rglob('__pycache__'))
    paths=[REPORT]+[ROOT/'docs/performance'/n for n in INDEX_NAMES]
    paths+=sorted(f for f in P.rglob('*') if f.is_file() and not local_only(f))
    if adopted: paths += [ROOT/name for name in SOURCE_NAMES]
    assert all(not f.is_symlink() for f in paths)
    preserved=load('environment.json')['preserved_files']
    for row in preserved:
        if row['path'] != '.git/index': assert sha(ROOT/row['path'])==row['sha256'],row['path']
    assert load('iwa-resolution-after.json')['status']=='pass'
    assert load('iwa-cleanup.json')['status']=='pass'
    index=Path(git('rev-parse','--git-path','index').decode().strip())
    if not index.is_absolute(): index=ROOT/index
    shared_index=sha(index); head=git('rev-parse','HEAD').decode().strip(); cherry=None; assert subprocess.run(['git','rev-parse','-q','--verify','CHERRY_PICK_HEAD'],cwd=ROOT,stdout=subprocess.DEVNULL).returncode!=0
    cleanup=load('cleanup.json')
    closure={'status':'pass','parent_commit':PARENT,'branch':BRANCH,'shared_head':head,'shared_cherry_pick':cherry,'shared_index_sha256':shared_index,'integration':'iWork cherry-pick resolved; performance branch integration pending','production_changes':adopted,'production_scope':'0913 performance candidate only','iwa_commit':load('iwa-resolution-after.json')['commit'],'decision':desc(P/'decision.json'),'native_children':48,'measured_samples':960,'guarded_operations':2400,'replay_before_and_after_cleanup':'pass','worktrees_created':1,'iwa_cleanup':load('iwa-cleanup.json'),'cleanup':cleanup,'local_only_evidence':[desc(f) for f in sorted(P.rglob('*')) if f.is_file() and local_only(f)],'files':[desc(f) for f in sorted(paths)]}
    with (P/'closure.json').open('x') as out:json.dump(closure,out,indent=2,sort_keys=True);out.write('\n')
    paths.append(P/'closure.json'); names=sorted(str(f.relative_to(ROOT)) for f in paths)
    with tempfile.TemporaryDirectory(prefix='litchi-0913-commit-index-') as directory:
        directory=Path(directory); env=dict(os.environ,GIT_INDEX_FILE=str(directory/'index'));git('read-tree',PARENT,env=env)
        specfile=directory/'paths';specfile.write_bytes(b'\0'.join(n.encode() for n in names)+b'\0')
        git('add','-f','--pathspec-from-file='+str(specfile),'--pathspec-file-nul',env=env)
        changed=sorted(git('diff','--cached','--name-only','-z',PARENT,env=env).decode().strip('\0').split('\0'));assert changed==names,sorted(set(changed)^set(names))
        authored=[n for n in names if not n.endswith(('.txt','.patch'))]
        git('-c','core.whitespace=trailing-space,space-before-tab,-blank-at-eof','diff','--cached','--check',PARENT,'--',*authored,env=env)
        tree=git('write-tree',env=env).decode().strip()
        title='perf: bypass redundant owned OPC member spans' if adopted else 'perf: evaluate and reject owned OPC direct-read candidate'
        message=title+'\n\nRetain matched native measurements, correctness and preservation gates,\nindependent statistical checks, replay and temporary-root cleanup.\n'
        commit=git('commit-tree',tree,'-p',PARENT,env=env,data=message.encode()).decode().strip()
        for row in closure['files']:
            data=git('show',commit+':'+row['path']);assert len(data)==row['bytes'] and hashlib.sha256(data).hexdigest()==row['sha256'],row['path']
        assert git('show',commit+':'+str((P/'closure.json').relative_to(ROOT)))==(P/'closure.json').read_bytes()
        assert sha(index)==shared_index and git('rev-parse','HEAD').decode().strip()==head
        assert subprocess.run(['git','rev-parse','-q','--verify','CHERRY_PICK_HEAD'],cwd=ROOT,stdout=subprocess.DEVNULL).returncode!=0
        git('update-ref',BRANCH,commit,'0'*40)
    assert sha(index)==shared_index
    print(json.dumps({'status':'pass','commit':commit,'branch':BRANCH,'files':len(names),'adopted':adopted,'shared_index_unchanged':True,'temporary_index_removed':True},indent=2))
if __name__=='__main__':main()
