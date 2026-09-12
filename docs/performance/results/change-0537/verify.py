"""Verify current source, historical replay, and an unapplied candidate draft."""
import json
from pathlib import Path
import subprocess
import analyze

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

def verify():
    before = {str(f.relative_to(HERE)): analyze.sha(f) for f in HERE.rglob('*') if f.is_file()}
    result = analyze.analyze()
    assert result == json.loads((HERE/'analysis.json').read_text())
    dependency = json.loads((HERE/'dependency-binding.json').read_text())
    assert analyze.sha(Path(dependency['path'])) == dependency['sha256']
    for name, digest in json.loads((HERE/'adr-manifest.json').read_text())['files'].items():
        assert analyze.sha(REPO/name) == digest
    patch = HERE/'candidate.patch'
    cleanup = json.loads((HERE/'cleanup.json').read_text())
    assert cleanup['removed'] and not Path(cleanup['path']).exists()
    assert cleanup['candidate_patch_sha256'] == analyze.sha(patch)
    subprocess.run(['git','apply','--check',str(patch)],cwd=REPO,check=True,capture_output=True)
    assert not subprocess.check_output(['git','diff','--name-only','--','crates','tools'],cwd=REPO)
    after = {str(f.relative_to(HERE)): analyze.sha(f) for f in HERE.rglob('*') if f.is_file()}
    assert before == after, 'verification modified evidence'
    if (HERE/'SHA256SUMS').exists():
        analyze.seal(HERE, analyze.sha(HERE/'SHA256SUMS'))
    return dict(status='pass', historical_rows=len(result['rows']),
                relevant_source_files=len(json.loads((HERE/'source-binding.json').read_text())),
                candidate_patch_sha256=analyze.sha(patch), draft_applies=True,
                runtime_changed=False, fresh_rust_tests_or_measurements=False,
                analysis_sha256=analyze.sha(HERE/'analysis.json'),
                verifier_sha256=analyze.sha(Path(__file__)))

if __name__ == '__main__':
    print(json.dumps(verify(),indent=2,sort_keys=True))
