"""Join the validated performance branch and resolved iWork branch without rewriting files."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
PERF='refs/heads/perf/0913-opc-owned-direct-read'
TARGET='refs/heads/feat/office-format-completeness'
def git(*args,env=None,data=None):return subprocess.check_output(['git',*args],cwd=ROOT,env=env,input=data)
def main():
    assert not (P/'integration.json').exists()
    before=git('rev-parse','HEAD').decode().strip();iwa=json.loads((P/'iwa-resolution-after.json').read_text());assert before==iwa['commit']
    assert git('symbolic-ref','HEAD').decode().strip()==TARGET
    assert not git('diff','--cached','--name-only')
    assert subprocess.run(['git','rev-parse','-q','--verify','CHERRY_PICK_HEAD'],cwd=ROOT,stdout=subprocess.DEVNULL).returncode!=0
    perf=git('rev-parse',PERF).decode().strip();base=git('merge-base',before,perf).decode().strip();assert base==iwa['parent']
    shared_index=ROOT/'.git/index'; index_hash=hashlib.sha256(shared_index.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix='litchi-0913-integration-index-') as directory:
        env=dict(os.environ,GIT_INDEX_FILE=str(Path(directory)/'index'))
        git('read-tree','-m',base,before,perf,env=env);assert not git('ls-files','-u',env=env)
        tree=git('write-tree',env=env).decode().strip()
        names=[n for n in git('diff','--name-only','--no-renames','-z',before,tree,env=env).decode().split('\0') if n]
        for name in names:
            data=git('show',tree+':'+name,env=env);assert (ROOT/name).is_file() and (ROOT/name).read_bytes()==data,name
        # Every existing tracked local modification must belong to this exact
        # verified merge result. Unrelated untracked files remain untouched.
        local=set(git('diff','--name-only','-z',before).decode().strip('\0').split('\0'))-{''}
        assert local<=set(names),sorted(local-set(names))
        result={'status':'pass','kind':'verified merge-tree integration','target':TARGET,'parents':[before,perf],'merge_base':base,'verified_existing_worktree_files':len(names),'worktree_file_contents_rewritten':False,'shared_index_preserved_until_publication':True}
        with (P/'integration.json').open('x') as out:json.dump(result,out,indent=2);out.write('\n')
        git('add','-f','--',str((P/'integration.json').relative_to(ROOT)),env=env)
        tree=git('write-tree',env=env).decode().strip()
        commit=git('commit-tree',tree,'-p',before,'-p',perf,env=env,data=b'Merge validated Office performance batches and resolved iWork ingress\n').decode().strip()
        index_lock=shared_index.with_name(shared_index.name+'.lock')
        # Hold Git's normal index lock across compare-and-swap publication, so
        # another process cannot stage entries between our checks and update.
        with index_lock.open('xb') as locked:
            try:
                locked.write(Path(env['GIT_INDEX_FILE']).read_bytes());locked.flush();os.fsync(locked.fileno())
                assert hashlib.sha256(shared_index.read_bytes()).hexdigest()==index_hash and git('rev-parse','HEAD').decode().strip()==before
                git('update-ref',TARGET,commit,before)
                os.replace(index_lock,shared_index)
            finally:
                if index_lock.exists():index_lock.unlink()
        assert git('rev-parse','HEAD').decode().strip()==commit and not git('diff','--cached','--name-only') and not git('diff','--name-only')
    print(json.dumps({'status':'pass','commit':commit,'performance_parent':perf,'iwa_parent':before,'target':TARGET,'temporary_index_removed':True},indent=2))
if __name__=='__main__':main()
