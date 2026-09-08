#!/usr/bin/env python3
"""Prepare a small clean checkout and authenticate every tracked Rust build source."""
import datetime
import hashlib
import json
from pathlib import Path
import shutil
import subprocess

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path('/tmp/litchi-goal-0474')
TREE = TEMP / 'tree'

def sha(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()

def write(path, value):
    with path.open('x') as stream:
        json.dump(value, stream, indent=2, sort_keys=True); stream.write('\n')

def main():
    revision = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip()
    tracked = subprocess.check_output(['git', 'ls-files', '-z'], cwd=REPO).decode().split('\0')
    tracked = [p for p in tracked if p]
    sources = {p: sha(REPO / p) for p in tracked if Path(p).suffix in ('.rs', '.toml', '.lock')}
    changed = subprocess.check_output(['git', 'diff', 'HEAD', '--name-only'], cwd=REPO, text=True).splitlines()
    assert not any(p in sources for p in changed), 'commit build inputs first'
    assert 'tools/perf-baseline/src/pptx_streaming_create.rs' in sources
    assert not TREE.exists()
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    subprocess.run(['git', 'worktree', 'add', '--detach', '--no-checkout', str(TREE), revision], cwd=REPO, check=True)
    selected = [p for p in tracked if not p.startswith('docs/') or p in sources]
    subprocess.run(['git', 'sparse-checkout', 'set', '--no-cone', '--stdin'], cwd=TREE,
                   input=''.join('/' + p + '\n' for p in selected), text=True, check=True)
    subprocess.run(['git', 'reset', '--hard', revision], cwd=TREE, check=True)
    fixtures = {
        'test-data/poi/test-data/spreadsheet/54016.xls': '2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a',
        'test-data/rtf/watermark.rtf': '48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb',
    }
    for path, digest in fixtures.items():
        assert sha(REPO / path) == digest
        target = TREE / path
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(REPO / path, target)
        assert sha(target) == digest
    for path, digest in sources.items():
        assert sha(TREE / path) == digest, path
    assert not subprocess.check_output(['git', 'status', '--porcelain'], cwd=TREE).strip()
    (ROOT / 'sources').mkdir(exist_ok=False)
    write(ROOT / 'sources/source.json', sources)
    binding = dict(revision=revision, source_manifest=dict(path='sources/source.json',
                   sha256=sha(ROOT / 'sources/source.json'), files=len(sources)), fixtures=fixtures)
    write(ROOT / 'source-binding.json', binding)
    write(ROOT / 'prepare.json', dict(status='pass', driver_sha256=sha(Path(__file__)),
          started_utc=started, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
          revision=revision, clean_after=True, build_path=str(TREE), selected_files=len(selected),
          selected_patterns_sha256=hashlib.sha256(''.join('/' + p + '\n' for p in selected).encode()).hexdigest(),
          source_binding_sha256=sha(ROOT / 'source-binding.json')))
    print('prepared', revision, len(sources), 'sources', flush=True)

if __name__ == '__main__':
    main()
