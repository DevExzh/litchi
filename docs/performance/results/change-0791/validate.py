"""Offline replay and final custody checks; never executes a native workload."""
import argparse
import os
import subprocess
import sys
import custody as c
import analysis_common as a

DOCS = ['0791-pptx-current-capture-profile.md', 'BASELINE.md', 'HOTSPOTS.md',
        'REPORT.md', 'CRUD_COVERAGE.md', 'GOAL_AUDIT.md']

def payloads():
    return {str(p.relative_to(c.P)): c.sha(p) for p in sorted(c.P.rglob('*'))
            if p.is_file() and p.name != 'seal.json' and '__pycache__' not in p.parts}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--require-final-seal', action='store_true')
    parser.add_argument('--check-workspace', action='store_true')
    args = parser.parse_args()
    for name in ['build/frozen-inputs.json', 'build/probe.json']:
        for path, digest in c.read(c.P/name).items():
            assert c.sha(c.P/path) == digest, path
    inheritance = c.read(c.P/'inheritance.json')
    for ref in [*inheritance['references'].values(), inheritance['production_reference']]:
        path = c.P/ref['path']
        assert path.stat().st_size == ref['bytes'] and c.sha(path) == ref['sha256']
    for path, digest in c.read(c.P/'architecture-inputs.json').items():
        assert c.sha(c.ROOT/path) == digest, path
    assert c.source()['files'] == c.read(c.P/'build/source.json')['files']
    quality = c.read(c.P/'quality.json')['format']
    assert quality['exit_code'] == 0
    a.artifact(quality['log'], 'format log')
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE='1')
    for script in ['native_analysis.py', 'profile_analysis.py', 'perf_analysis.py',
                   'root_scan_costs.py', 'root_frames.py']:
        subprocess.run([sys.executable, '-B', str(c.P/script), '--check'], check=True, env=env)
    if args.check_workspace:
        origin = c.read(c.P/'origin.json')
        for path, digest in origin['unrelated'].items():
            assert c.sha(c.ROOT/path) == digest, path
        trees = subprocess.check_output(['git', 'worktree', 'list', '--porcelain'], cwd=c.ROOT, text=True)
        assert trees.split('\n\n')[1:] == origin['worktrees'].split('\n\n')[1:]
    seal_path = c.P/'seal.json'
    if args.require_final_seal or seal_path.exists():
        seal = c.read(seal_path)
        assert payloads() == seal['files'], 'packet seal differs'
        assert seal['documents'] == {name:c.sha(c.P.parents[1]/name) for name in DOCS}
        assert not c.TARGET.exists()
        assert not list(c.P.rglob('__pycache__'))
        assert not any(p.is_symlink() for p in c.P.rglob('*'))
    print('0791 offline replay and custody PASS')

if __name__ == '__main__':
    main()
