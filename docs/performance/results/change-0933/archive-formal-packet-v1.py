#!/usr/bin/env python3
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import tarfile


def normalize(member):
    member.uid = 0
    member.gid = 0
    member.uname = ''
    member.gname = ''
    member.mtime = 0
    return member


def make_archive(packet, archive):
    packet = packet.resolve()
    archive = archive.resolve()
    if not packet.is_dir():
        raise SystemExit(f'packet directory missing: {packet}')
    if archive.exists():
        raise SystemExit(f'refusing to replace archive: {archive}')
    with archive.open('xb') as raw:
        with gzip.GzipFile(fileobj=raw, mode='wb', filename='', mtime=0) as compressed:
            with tarfile.open(fileobj=compressed, mode='w') as tar:
                tar.addfile(normalize(tar.gettarinfo(str(packet), arcname=packet.name)))
                for path in sorted(packet.rglob('*'), key=lambda item: str(item.relative_to(packet))):
                    arcname = f'{packet.name}/{path.relative_to(packet)}'
                    info = normalize(tar.gettarinfo(str(path), arcname=arcname))
                    if info.isreg():
                        with path.open('rb') as stream:
                            tar.addfile(info, stream)
                    else:
                        tar.addfile(info)


def descriptor(path):
    return {
        'path': str(path),
        'bytes': path.stat().st_size,
        'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('packet', type=Path)
    parser.add_argument('archive', type=Path)
    args = parser.parse_args()
    make_archive(args.packet, args.archive)
    print(json.dumps(descriptor(args.archive), sort_keys=True))
