#!/usr/bin/env python3
"""Check formatting for workspace packages within the user's non-iWork scope."""
import json
import subprocess

metadata = json.loads(subprocess.check_output([
    'cargo', 'metadata', '--locked', '--no-deps', '--format-version', '1',
]))
members = set(metadata['workspace_members'])
names = sorted(package['name'] for package in metadata['packages']
               if package['id'] in members
               and not package['name'].startswith(('litchi-iwa', 'litchi-numbers'))
               and package['name'] not in {'litchi-keynote', 'litchi-pages'})
assert names
argv = ['cargo', 'fmt']
for name in names:
    argv.extend(['-p', name])
argv.extend(['--', '--check'])
print(json.dumps({'packages': names, 'argv': argv}), flush=True)
subprocess.run(argv, check=True)
