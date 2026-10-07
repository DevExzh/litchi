#!/usr/bin/env python3
"""Create-only 0953 external archive helper; requires explicit --seal --go."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
import shutil
import stat
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

PACKET_DEFAULT = Path('/home/zhuhe/litchi-performance-0953')
ARCHIVE_DEFAULT = Path('/home/zhuhe/litchi-performance-0953-archive')
BINARY_DEFAULT = PACKET_DEFAULT / 'target-frame-pointer-native-v1/release/litchi-perf-baseline'
BINARY_RECEIPT_DEFAULT = PACKET_DEFAULT / 'builds/frame-pointer-native-v1/binary.json'
EXCLUDE_TOP = {'target', 'tmp'}
EXCLUDE_DIR_NAMES = {'__pycache__'}

def digest(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f'not a regular file: {path}')
    st = path.stat(); h = hashlib.sha256()
    size = 0
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1 << 20), b''):
            h.update(block); size += len(block)
    return {'path': str(path), 'bytes': size, 'mode': stat.S_IMODE(st.st_mode),
            'executable': bool(st.st_mode & 0o111), 'sha256': h.hexdigest()}

def excluded(rel: Path) -> bool:
    return bool(rel.parts and (rel.parts[0] in EXCLUDE_TOP or rel.parts[0].startswith('target-')))

def collect(packet: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    if not packet.is_dir() or packet.is_symlink():
        raise RuntimeError(f'invalid packet root: {packet}')
    entries: list[dict[str, Any]] = []
    directories: list[dict[str, Any]] = []
    excluded_paths: list[str] = []
    for current, dirnames, filenames in os.walk(packet, topdown=True, followlinks=False):
        current_path = Path(current)
        current_rel = current_path.relative_to(packet)
        kept: list[str] = []
        for name in sorted(dirnames):
            path = current_path / name
            rel = path.relative_to(packet)
            rel_text = rel.as_posix()
            if path.is_symlink():
                raise RuntimeError(f'symlink directory: {rel_text}')
            if excluded(rel):
                excluded_paths.append(rel_text + '/')
                continue
            if name in EXCLUDE_DIR_NAMES:
                excluded_paths.append(rel_text + '/')
                continue
            kept.append(name)
            directories.append({'path': rel_text, 'mode': stat.S_IMODE(path.stat().st_mode)})
        dirnames[:] = kept
        for name in sorted(filenames):
            path = current_path / name
            rel = path.relative_to(packet)
            rel_text = rel.as_posix()
            if path.is_symlink():
                raise RuntimeError(f'symlink file: {rel_text}')
            if name.endswith('.pyc'):
                excluded_paths.append(rel_text)
                continue
            if excluded(rel):
                excluded_paths.append(rel_text)
                continue
            entries.append({**digest(path), 'path': rel_text})
    entries.sort(key=lambda item: str(item['path']))
    directories.sort(key=lambda item: str(item['path']))
    excluded_paths.sort()
    return entries, directories, excluded_paths

def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except Exception as exc:
        raise RuntimeError(f'invalid JSON {path}: {exc}') from exc

def retain_binary(source: Path, receipt_path: Path, destination: Path) -> dict[str, Any]:
    source = source.resolve(); receipt_path = receipt_path.resolve(); destination = destination.resolve()
    if destination.exists() or destination.is_symlink():
        raise RuntimeError(f'refusing existing binary destination: {destination}')
    actual = digest(source)
    receipt = load_json(receipt_path)
    bound = receipt.get('binary', {})
    if bound.get('status') != 'present' or bound.get('sha256') != actual['sha256'] or bound.get('bytes') != actual['bytes']:
        raise RuntimeError('binary receipt does not bind the source binary')
    if not actual['executable']:
        raise RuntimeError('profiling binary is not executable')
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, destination)
    os.chmod(destination, actual['mode'])
    copied = digest(destination)
    if copied['sha256'] != actual['sha256'] or copied['bytes'] != actual['bytes']:
        raise RuntimeError('binary role copy differs')
    return {'source': actual, 'destination': copied, 'receipt': digest(receipt_path)}

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--packet', type=Path, default=PACKET_DEFAULT)
    ap.add_argument('--archive-root', type=Path, default=ARCHIVE_DEFAULT)
    ap.add_argument('--binary', type=Path, default=BINARY_DEFAULT)
    ap.add_argument('--binary-receipt', type=Path, default=BINARY_RECEIPT_DEFAULT)
    ap.add_argument('--seal', action='store_true')
    ap.add_argument('--go', action='store_true')
    args = ap.parse_args()
    if args.seal and not args.go:
        raise SystemExit('--seal requires --go')
    packet = args.packet.resolve(); root = args.archive_root.resolve()
    root.mkdir(parents=True, exist_ok=True)
    entries, directories, excluded_paths = collect(packet)
    manifest = {
      'schema':'litchi.performance.0953.external-archive-manifest.v1',
      'status':'sealed' if args.seal else 'planning_only_no_archive',
      'generated_utc':datetime.now(timezone.utc).isoformat(),
      'packet_root':str(packet), 'entry_count':len(entries),
      'logical_bytes':sum(int(x['bytes']) for x in entries),
      'entries':entries, 'directories':directories, 'excluded':excluded_paths,
      'symlinks':[], 'dependencies':{
        '0950':{'path':'/home/zhuhe/litchi-performance-0950','copied':False},
        '0952':{'path':'/home/zhuhe/litchi-performance-0952','copied':False}},
      'cleanup_requires_explicit_go':True,
      'binary_external_to_core_archive':True,
    }
    manifest_path = root/'manifest-v1.json'
    if args.seal and manifest_path.exists():
        raise RuntimeError(f'refusing existing sealed manifest: {manifest_path}')
    manifest_path.write_text(json.dumps(manifest, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    if not args.seal:
        print(json.dumps({'status':manifest['status'],'entries':len(entries),'logical_bytes':manifest['logical_bytes']}))
        return 0
    for name in ('selection-v1.list','archive-0953-v1.tar','archive-0953-v1.tar.zst','archive-receipt-v1.json'):
        if (root/name).exists(): raise RuntimeError(f'refusing existing archive artifact: {root/name}')
    selection=root/'selection-v1.list'
    with selection.open('wb') as stream:
        for item in directories+entries:
            stream.write(str(item['path']).encode()+b'\0')
    tar_path=root/'archive-0953-v1.tar'; zst_path=root/'archive-0953-v1.tar.zst'; tmp=root/'.archive-0953-v1.tar.zst.tmp'
    subprocess.run(['/usr/bin/tar','--create','--no-recursion','--sort=name','--mtime=@0','--owner=0','--group=0','--numeric-owner','--format=pax','--file',str(tar_path),'--null','--verbatim-files-from','--files-from',str(selection)],cwd=packet,check=True)
    subprocess.run(['/usr/bin/zstd','-3','--no-progress','--force',str(tar_path),'-o',str(tmp)],check=True)
    tmp.replace(zst_path)
    tar_desc=digest(tar_path); tar_path.unlink(); archive_desc=digest(zst_path)
    binary = retain_binary(args.binary, args.binary_receipt, root/'binaries/litchi-perf-baseline-frame-pointer-v1')
    receipt={'schema':'litchi.performance.0953.external-archive-receipt.v1','status':'sealed','manifest':digest(manifest_path),'selection':digest(selection),'compressed_archive':archive_desc,'uncompressed_tar':{**tar_desc,'deleted_after_compression':True},'helper':digest(Path(__file__)),'packet_census':{'entries':len(entries),'logical_bytes':manifest['logical_bytes'],'directories':len(directories),'excluded':excluded_paths,'symlinks':[]},'binary_external':binary,'dependencies':manifest['dependencies'],'regeneration':f'python3 -B {Path(__file__).resolve()} --packet {packet} --archive-root <fresh-archive-root> --seal --go'}
    (root/'archive-receipt-v1.json').write_text(json.dumps(receipt,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print(json.dumps({'status':'sealed','entries':len(entries),'logical_bytes':manifest['logical_bytes'],'archive_bytes':archive_desc['bytes'],'archive_sha256':archive_desc['sha256'],'binary_sha256':binary['destination']['sha256']}))
    return 0
if __name__=='__main__': raise SystemExit(main())
