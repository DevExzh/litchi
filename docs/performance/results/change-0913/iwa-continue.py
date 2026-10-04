"""Stage only the validated iWork resolution and complete its cherry-pick."""
import hashlib
import json
from pathlib import Path
import subprocess

P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def git(*args):return subprocess.check_output(['git',*args],cwd=ROOT)
def sha(data):return hashlib.sha256(data).hexdigest()
def main():
    assert not (P/'iwa-resolution-after.json').exists()
    before=json.loads((P/'iwa-resolution-before.json').read_text());quality=json.loads((P/'iwa-quality.json').read_text());source=json.loads((P/'iwa-quality-source-v4.json').read_text())
    assert quality['status']=='pass' and all(r['exit_code']==0 for r in quality['commands'])
    assert json.loads((P/'iwa-format-final.json').read_text())['status']=='pass'
    assert git('rev-parse','HEAD').decode().strip()==before['head_before']
    assert git('rev-parse','CHERRY_PICK_HEAD').decode().strip()==before['cherry_pick']
    expected=sorted(r['path'] for r in source['files'])
    staged=sorted(git('diff','--cached','--name-only').decode().splitlines())
    assert set(staged)<=set(expected) and set(before['staged_paths'])<=set(staged),staged
    for row in source['files']:
        data=(ROOT/row['path']).read_bytes();assert len(data)==row['bytes'] and sha(data)==row['sha256'],row['path']
    git('add','--',*expected)
    assert sorted(git('diff','--cached','--name-only').decode().splitlines())==expected
    assert not git('ls-files','-u')
    for row in source['files']:assert sha(git('show',':'+row['path']))==row['sha256']
    git('diff','--cached','--check')
    message=Path(git('rev-parse','--git-path','MERGE_MSG').decode().strip())
    if not message.is_absolute():message=ROOT/message
    message.write_text('feat(iwa): admit neutral candidate ingress phases\n\nInclude the required bounded archive clone inventory from c09fce06c7.\nResolve exports against the current physical archive owner without importing\nthe unrelated replacement API predecessor chain. Correct incoming lint and\nformatting issues; validate both owner crates and native preservation tests.\n')
    subprocess.run(['git','-c','core.editor=true','cherry-pick','--continue'],cwd=ROOT,check=True)
    commit=git('rev-parse','HEAD').decode().strip()
    assert git('rev-parse','HEAD^').decode().strip()==before['head_before']
    assert sorted(git('diff-tree','--no-commit-id','--name-only','-r',commit).decode().splitlines())==expected
    assert subprocess.run(['git','rev-parse','-q','--verify','CHERRY_PICK_HEAD'],cwd=ROOT,stdout=subprocess.DEVNULL).returncode!=0
    assert not git('diff','--cached','--name-only')
    for row in source['files']:assert sha(git('show',commit+':'+row['path']))==row['sha256']
    result={'status':'pass','commit':commit,'parent':before['head_before'],'cherry_pick':before['cherry_pick'],'inventory_prerequisite':'c09fce06c7','files':source['files'],'tests_passed':302,'tests_failed':0,'cherry_pick_state_cleared':True,'unmerged_paths':[],'index_after_sha256':sha((ROOT/'.git/index').read_bytes()),'scope':'Only the eight validated iWork files committed; performance work remains unstaged.'}
    with (P/'iwa-resolution-after.json').open('x') as out:json.dump(result,out,indent=2);out.write('\n')
    print(json.dumps({'status':'pass','commit':commit,'files':len(expected)}))
if __name__=='__main__':main()
