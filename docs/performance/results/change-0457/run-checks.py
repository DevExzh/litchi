#!/usr/bin/env python3
"""Run serialized validation gates for bounded ZIP/ODF/ODP publication."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
COMMANDS = {
    'format': ['cargo', 'fmt', '--all', '--', '--check'],
    'harness-format': ['cargo', 'fmt', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--', '--check'],
    'zip': ['cargo', 'test', '--locked', '--release', '-p', 'soapberry-zip', '--all-features', '--', '--test-threads=1'],
    'odf-common': ['cargo', 'test', '--locked', '--release', '-p', 'litchi-odf-common', '--all-features', '--', '--test-threads=1'],
    'odp': ['cargo', 'test', '--locked', '--release', '-p', 'litchi-odp', '--all-features', '--', '--test-threads=1'],
    'opc': ['cargo', 'test', '--locked', '--release', '-p', 'litchi-opc', '--all-features', '--', '--test-threads=1'],
    'harness': ['cargo', 'test', '--locked', '--release', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--all-features', '--', '--test-threads=1'],
    'strict': ['cargo', 'clippy', '--locked', '--release', '-p', 'soapberry-zip', '-p', 'litchi-odf-common', '-p', 'litchi-odp', '--all-targets', '--all-features', '--', '-D', 'warnings'],
    'harness-strict': ['cargo', 'clippy', '--locked', '--release', '--manifest-path', 'tools/perf-baseline/Cargo.toml', '--all-targets', '--all-features', '--', '-D', 'warnings'],
    'doc': ['cargo', 'doc', '--locked', '--release', '-p', 'soapberry-zip', '-p', 'litchi-odf-common', '-p', 'litchi-odp', '--no-deps', '--all-features'],
    'workspace': ['cargo', 'check', '--locked', '--release', '--workspace', '--all-targets', '--no-default-features', '--exclude', 'litchi-iwa*', '--exclude', 'litchi-keynote', '--exclude', 'litchi-numbers*', '--exclude', 'litchi-pages', '--features', 'litchi/odf'],
    'boundaries': ['python3', '-B', 'tools/check_crate_boundaries.py'],
}

if __name__ == '__main__':
    for tag in sys.argv[1:] or COMMANDS:
        subprocess.run([sys.executable, '-B', str(ROOT / 'check.py'), '--tag', tag, '--', *COMMANDS[tag]], check=True)
