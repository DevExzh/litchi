"""Verify source, retained evidence, unapplied draft and cleanup for audit0526."""
import argparse
import ast
import hashlib
import json
import os
import tempfile
from pathlib import Path
import subprocess

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PRIOR = HERE.parent / 'change-0525'
EXPECTED_FILES = {
    'crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs',
    'crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs',
    'crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs',
}

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def require(value, message):
    if not value:
        raise ValueError(message)

def inventory(root):
    require(not any(p.is_symlink() for p in root.rglob('*')), 'symlink in evidence')
    return {p.relative_to(root).as_posix(): sha(p) for p in root.rglob('*')
            if p.is_file() and p.name != 'SHA256SUMS'}

def check_seal(root):
    expected = {}
    for line in (root / 'SHA256SUMS').read_text().splitlines():
        digest, name = line.split('  ', 1)
        require(not Path(name).is_absolute() and '..' not in Path(name).parts
                and name not in expected and name != 'SHA256SUMS', 'unsafe/duplicate seal path')
        expected[name] = digest
    require(expected == inventory(root), 'seal inventory mismatch')
    return len(expected)

def replay_candidate(revision):
    with tempfile.TemporaryDirectory(prefix='litchi-0526-index-') as temp:
        env = dict(os.environ, GIT_INDEX_FILE=str(Path(temp) / 'index'))
        subprocess.run(['git', 'read-tree', revision], cwd=REPO, env=env, check=True)
        subprocess.run(['git', 'apply', '--cached', str(HERE / 'row-primary-arena.patch')],
                       cwd=REPO, env=env, check=True)
        return {name: hashlib.sha256(subprocess.check_output(
            ['git', 'show', ':' + name], cwd=REPO, env=env)).hexdigest()
            for name in sorted(EXPECTED_FILES)}

def verify(sealed=False):
    binding = json.loads((HERE / 'source-binding.json').read_text())
    for section in ('source_files', 'adr_files'):
        for name, digest in binding[section].items():
            require(sha(REPO / name) == digest, f'current source binding differs: {name}')
            original = subprocess.check_output(['git', 'show', f"{binding['revision']}:{name}"], cwd=REPO)
            require(hashlib.sha256(original).hexdigest() == digest, f'base source binding differs: {name}')
    for name, item in binding.get('working_only_inputs', {}).items():
        require(sha(REPO / name) == item['sha256'] == sha(HERE / item['retained']),
                f'working-only input differs: {name}')
    prior = json.loads((HERE / 'prior-seal-verification.json').read_text())
    require(sha(PRIOR / 'SHA256SUMS') == prior['seal_sha256'], 'prior seal binding differs')
    prior_count = check_seal(PRIOR)
    require(prior_count == prior['exact_inventory_entries'] == 550, 'prior inventory size differs')
    patch = HERE / 'row-primary-arena.patch'
    paths = {line.split(' b/', 1)[1] for line in patch.read_text().splitlines()
             if line.startswith('diff --git a/')}
    require(paths == EXPECTED_FILES, 'draft patch scope differs')
    result = subprocess.run(['git', 'apply', '--check', str(patch)], cwd=REPO, text=True, capture_output=True)
    require(result.returncode == 0, f'draft patch does not apply: {result.stderr}')
    candidate = json.loads((HERE / 'candidate-source-binding.json').read_text())
    require(candidate['patch_sha256'] == sha(patch), 'candidate patch binding differs')
    require(candidate['files'] == replay_candidate(binding['revision']), 'candidate source replay differs')
    for script in HERE.glob('*.py'):
        ast.parse(script.read_text(), filename=str(script))
    with tempfile.TemporaryDirectory(prefix='litchi-0526-profile-replay-') as temp:
        output = Path(temp) / 'profile-analysis.json'
        subprocess.run(['python3', '-B', str(HERE / 'analyze.py'), '--output', str(output)],
                       cwd=REPO, check=True)
        require(output.read_bytes() == (HERE / 'profile-analysis.json').read_bytes(),
                'retained profile decomposition does not replay')
    require(not Path('/tmp/litchi-goal-0526-candidate').exists(), 'candidate scratch remains')
    require(not list(HERE.rglob('__pycache__')), 'Python cache remains')
    decision = json.loads((HERE / 'decision.json').read_text())
    require(decision['status'] == 'audit_complete_draft_ready'
            and decision['production_change_retained'] is False
            and decision['new_performance_capture'] is False, 'audit disposition differs')
    for name, digest in decision['bindings'].items():
        require(not Path(name).is_absolute() and '..' not in Path(name).parts
                and sha(HERE / name) == digest, f'decision artifact binding differs: {name}')
    return {'status': 'pass', 'source_revision': binding['revision'],
            'source_files': len(binding['source_files']), 'adr_files': len(binding['adr_files']),
            'prior_seal_entries': prior_count, 'draft_patch_sha256': sha(patch),
            'draft_applies': True, 'candidate_source_replay': True, 'profile_analysis_replay': True, 'candidate_scratch_absent': True,
            'python_cache_absent': True, 'seal_entries': check_seal(HERE) if sealed else None,
            'scope': 'Source/ADR/retained seal/draft applicability/cleanup audit; not a Rust build, correctness suite or new performance measurement.'}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--sealed', action='store_true')
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = verify(args.sealed)
    text = json.dumps(report, indent=2) + '\n'
    if args.output:
        require(not args.sealed, 'sealed verification must not mutate its own inventory')
        args.output.write_text(text)
    print(text, end='')
