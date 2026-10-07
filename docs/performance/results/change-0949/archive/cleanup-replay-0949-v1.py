#!/usr/bin/env python3
"""Remove only completed 0949 replay/extraction scratch and generated pycache."""
from __future__ import annotations
import hashlib, json, os, shutil, stat
from datetime import datetime, timezone
from pathlib import Path

ARCHIVE = Path('/home/zhuhe/litchi-performance-0949-archive')
RECEIPT = ARCHIVE / 'replay-cleanup-receipt-v1.json'
ROOTS = [
    Path('/home/zhuhe/litchi-performance-0949-replay-v1-20261007'),
    Path('/home/zhuhe/litchi-performance-0949-offline-replay-v1-20261007'),
    Path('/home/zhuhe/litchi-performance-0949/tooling/__pycache__'),
]

def sha(path: Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest()

def inventory(root: Path) -> dict:
    if not root.exists(): return {'exists':False,'files':0,'directories':0,'bytes':0,'symlinks':0}
    if not root.is_dir() or root.is_symlink(): raise SystemExit(f'not a real directory: {root}')
    files=dirs=bytes_=symlinks=0
    for cur,dnames,fnames in os.walk(root,topdown=True,followlinks=False):
        dirs += len(dnames)
        for n in dnames:
            if (Path(cur)/n).is_symlink(): symlinks += 1
        for n in fnames:
            p=Path(cur)/n
            if p.is_symlink(): symlinks += 1
            else: files += 1; bytes_ += p.stat().st_size
    return {'exists':True,'files':files,'directories':dirs,'bytes':bytes_,'symlinks':symlinks}

def main():
    before={str(p):inventory(p) for p in ROOTS}
    for p,i in before.items():
        if not i['exists']: raise SystemExit(f'missing owned cleanup root: {p}')
        if i['symlinks']: raise SystemExit(f'symlink in cleanup root: {p}')
    data={
      'schema':'litchi.performance.0949.replay_cleanup.v1',
      'status':'started',
      'started_utc':datetime.now(timezone.utc).isoformat(),
      'roots':[str(p) for p in ROOTS],
      'before':before,
      'preserved_receipts':[
        str(ARCHIVE/'archive-receipt-v1.json'),
        str(ARCHIVE/'replay-v1-extraction-verification.json'),
        str(ARCHIVE/'offline-replay-v1-failure-receipt.json'),
        str(ARCHIVE/'offline-replay-v2-validation-receipt.json'),
      ],
      'source_packet_retained':True,
      'archive_retained':True,
    }
    RECEIPT.write_text(json.dumps(data,indent=2,sort_keys=True)+'\n')
    for p in ROOTS: shutil.rmtree(p)
    after={str(p):inventory(p) for p in ROOTS}
    if any(i['exists'] for i in after.values()): raise SystemExit('cleanup verification failed')
    data.update({'status':'pass','completed_utc':datetime.now(timezone.utc).isoformat(),'after':after,
                 'removed_files':sum(i['files'] for i in before.values()),
                 'removed_directories':sum(i['directories'] for i in before.values()),
                 'removed_logical_bytes':sum(i['bytes'] for i in before.values()),
                 'helper_sha256':sha(Path(__file__).resolve())})
    RECEIPT.write_text(json.dumps(data,indent=2,sort_keys=True)+'\n')
    print(json.dumps({k:data[k] for k in ('status','removed_files','removed_directories','removed_logical_bytes','helper_sha256')},sort_keys=True))

if __name__=='__main__': main()
