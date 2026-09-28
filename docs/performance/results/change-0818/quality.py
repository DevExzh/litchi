"""Root-only serial affected-crate quality gates."""
from pathlib import Path
import subprocess
import sys

P = Path(__file__).resolve().parent
steps = [
    ('fmt', ['cargo', 'fmt', '-p', 'litchi-docx', '--', '--check']),
    ('check', ['cargo', 'check', '--offline', '--locked', '-p', 'litchi-docx', '--all-features', '--all-targets']),
    ('tests', ['cargo', 'test', '--offline', '--locked', '-p', 'litchi-docx', '--all-features', '--', '--test-threads=2']),
    ('clippy', ['cargo', 'clippy', '--offline', '--locked', '-p', 'litchi-docx', '--all-features', '--all-targets', '--', '-D', 'warnings']),
    ('rustdoc', ['cargo', 'rustdoc', '--offline', '--locked', '-p', 'litchi-docx', '--all-features', '--lib', '--', '-D', 'warnings']),
    ('boundaries', ['python3', '-B', 'tools/check_crate_boundaries.py']),
]
for label, command in steps:
    result = subprocess.run([sys.executable, '-B', str(P / 'run.py'), label, 'quality', *command])
    if result.returncode:
        raise SystemExit(result.returncode)
print('0818 six quality gates complete', flush=True)
