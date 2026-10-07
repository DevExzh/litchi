#!/usr/bin/env python3
"""Archive a regular-file evidence packet and verify its exact extracted bytes."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess


def digest(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def inventory(root):
    rows = []
    for path in sorted(root.rglob('*')):
        if path.is_symlink():
            raise ValueError(f'symlink in packet: {path}')
        if path.is_dir():
            continue
        if not path.is_file():
            raise ValueError(f'nonregular packet member: {path}')
        rows.append({'path': path.relative_to(root).as_posix(),
                     'bytes': path.stat().st_size, 'sha256': digest(path)})
    return rows


def write_once(path, obj):
    with path.open('x') as stream:
        json.dump(obj, stream, indent=2, sort_keys=True)
        stream.write('\n')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--packet', required=True, type=Path)
    parser.add_argument('--destination', required=True, type=Path)
    args = parser.parse_args()
    packet = args.packet.resolve(strict=True)
    dest = args.destination.resolve()
    if dest == packet or packet in dest.parents:
        raise ValueError('destination must be outside packet')
    dest.mkdir(parents=True, exist_ok=True)
    archive = dest / 'litchi-performance-0946-v1.tar.gz'
    manifest = dest / 'inventory-v1.json'
    member_list = dest / 'members-v1.list'
    receipt = dest / 'archive-verified-v1.json'
    extraction = dest / 'verified-extraction'
    for path in (archive, manifest, member_list, receipt, extraction):
        if path.exists() or path.is_symlink():
            raise ValueError(f'refusing replacement: {path}')
    rows = inventory(packet)
    write_once(manifest, {'files': rows, 'count': len(rows),
                          'bytes': sum(row['bytes'] for row in rows)})
    with member_list.open('xb') as stream:
        for row in rows:
            stream.write((packet.name + '/' + row['path']).encode() + b'\0')
    command = ['/usr/bin/tar', '--create', '--gzip', '--format=posix',
               '--numeric-owner', '--owner=0', '--group=0', '--mtime=@0',
               '--pax-option=delete=atime,delete=ctime', '--no-recursion',
               '--null', '--verbatim-files-from', '--directory', str(packet.parent),
               '--files-from', str(member_list), '--file', str(archive)]
    subprocess.run(command, check=True)
    if inventory(packet) != rows:
        raise ValueError('packet changed during archive creation')
    extraction.mkdir()
    extract_command = ['/usr/bin/tar', '--extract', '--gzip', '--file', str(archive),
                       '--directory', str(extraction), '--no-same-owner']
    subprocess.run(extract_command, check=True)
    if inventory(extraction / packet.name) != rows:
        raise ValueError('extracted file set or hashes differ')
    record = {'schema': 'litchi.performance.0946.archive.v1', 'status': 'pass',
              'packet': str(packet), 'archive': {'path': str(archive),
              'bytes': archive.stat().st_size, 'sha256': digest(archive)},
              'inventory': {'path': str(manifest), 'bytes': manifest.stat().st_size,
                            'sha256': digest(manifest)},
              'files': len(rows), 'uncompressed_bytes': sum(r['bytes'] for r in rows),
              'exact_extracted_file_set_and_hashes': True,
              'archive_command': command, 'extract_command': extract_command,
              'extraction': str(extraction),
              'script_sha256': digest(Path(__file__)),
              'regeneration': ['python3', '-B', str(Path(__file__).resolve()),
                               '--packet', str(packet), '--destination', 'FRESH_EXTERNAL_DIRECTORY']}
    write_once(receipt, record)
    print(json.dumps({'status': 'pass', 'files': len(rows), 'archive': record['archive']}))


if __name__ == '__main__':
    main()
