"""Seal retained evidence, and audit every expected staged or committed blob."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
DOCS = ['0796-xml-attribute-boundary-diagnostic.md', 'BASELINE.md',
        'CRUD_COVERAGE.md', 'GOAL_AUDIT.md', 'HOTSPOTS.md', 'REPORT.md']

def digest(data):
    return hashlib.sha256(data).hexdigest()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--write', action='store_true')
    parser.add_argument('--check-index', action='store_true')
    parser.add_argument('--check-head', action='store_true')
    args = parser.parse_args()
    files = {str(f.relative_to(P)): digest(f.read_bytes()) for f in sorted(P.rglob('*'))
             if f.is_file() and f.name != 'seal.json' and '__pycache__' not in f.parts}
    assert not list(P.rglob('__pycache__'))
    docs = {'docs/performance/' + name: digest((ROOT/'docs/performance'/name).read_bytes())
            for name in DOCS}
    baseline = json.loads((P/'source.json').read_text())['files']
    production = {name: digest((ROOT/name).read_bytes()) for name, sha in baseline.items() if digest((ROOT/name).read_bytes()) != sha}
    assert not production, '0796 is diagnostic-only and must restore all production source'
    seal = {'production': production, 'schema': 'litchi.performance.0796.final-seal.v1', 'files': files, 'documents': docs}
    encoded = json.dumps(seal, indent=2, sort_keys=True) + '\n'
    if args.write:
        (P/'seal.json').write_text(encoded)
    assert (P/'seal.json').read_text() == encoded
    expected = {str(P.relative_to(ROOT)/name): value for name,value in files.items()}
    expected.update(docs)
    expected.update(production)
    expected[str((P/'seal.json').relative_to(ROOT))] = digest(encoded.encode())
    if args.check_index or args.check_head:
        rev = ':' if args.check_index else 'HEAD:'
        listing = ['git', 'diff', '--cached', '--name-only', '-z'] if args.check_index else ['git', 'diff-tree', '--no-commit-id', '--name-only', '-r', '-z', 'HEAD']
        changed = set(filter(None, subprocess.check_output(listing, cwd=ROOT).decode().split('\0')))
        assert changed == set(expected), (changed-set(expected), set(expected)-changed)
        proc = subprocess.Popen(['git','cat-file','--batch'], cwd=ROOT, stdin=subprocess.PIPE, stdout=subprocess.PIPE)
        for name, sha in expected.items():
            proc.stdin.write((rev+name+'\n').encode()); proc.stdin.flush()
            header = proc.stdout.readline().decode().split()
            assert len(header) == 3 and header[1] == 'blob', (name, header)
            data = proc.stdout.read(int(header[2])); assert proc.stdout.read(1) == b'\n'
            assert digest(data) == sha, name
        proc.stdin.close(); assert proc.wait() == 0
    print(f'0796 seal audit PASS: {len(files)} payloads + seal + {len(docs)} documents + {len(production)} production = {len(expected)} blobs')

if __name__ == '__main__':
    main()
