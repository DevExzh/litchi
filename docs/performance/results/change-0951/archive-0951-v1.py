#!/usr/bin/env python3
"""Seal the completed 0951 evidence packet; requires explicit --seal --go."""
from __future__ import annotations
import argparse, hashlib, json, os, stat, subprocess
from datetime import datetime, timezone
from pathlib import Path

PACKET_DEFAULT=Path('/home/zhuhe/litchi-performance-0951')
ARCHIVE_DEFAULT=Path('/home/zhuhe/litchi-performance-0951-archive')
SCHEMA='litchi.performance.0951.external-archive-manifest.v1'
EXCLUDE_DIRS={'__pycache__'}

def digest(path:Path):
    h=hashlib.sha256(); n=0
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1<<20),b''):
            h.update(b); n += len(b)
    st=path.stat()
    return {'path':path.as_posix(),'bytes':n,'mode':stat.S_IMODE(st.st_mode),'executable':bool(st.st_mode & 0o111),'sha256':h.hexdigest()}

def collect(packet:Path):
    if not packet.is_dir() or packet.is_symlink(): raise RuntimeError('invalid packet root')
    entries=[]; dirs=[]; excluded=[]
    for cur,dnames,fnames in os.walk(packet,topdown=True,followlinks=False):
        curp=Path(cur); keep=[]
        for name in sorted(dnames):
            p=curp/name; rel=p.relative_to(packet).as_posix()
            if p.is_symlink(): raise RuntimeError(f'symlink directory: {rel}')
            if name in EXCLUDE_DIRS:
                excluded.append(rel+'/'); continue
            keep.append(name)
        dnames[:]=keep
        for name in sorted(fnames):
            p=curp/name; rel=p.relative_to(packet).as_posix()
            if p.is_symlink(): raise RuntimeError(f'symlink file: {rel}')
            if p.suffix=='.pyc': excluded.append(rel); continue
            d=digest(p); d['path']=rel; entries.append(d)
        relcur='' if curp==packet else curp.relative_to(packet).as_posix()
        if relcur and not keep and not any((curp/n).is_file() for n in fnames):
            dirs.append({'path':relcur,'mode':stat.S_IMODE(curp.stat().st_mode)})
    entries.sort(key=lambda x:x['path']); dirs.sort(key=lambda x:x['path']); excluded.sort()
    return entries,dirs,excluded

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--packet',type=Path,default=PACKET_DEFAULT); ap.add_argument('--archive-root',type=Path,default=ARCHIVE_DEFAULT); ap.add_argument('--seal',action='store_true'); ap.add_argument('--go',action='store_true'); ap.add_argument('--manifest-only',action='store_true')
    a=ap.parse_args(); packet=a.packet.resolve(); root=a.archive_root.resolve(); root.mkdir(parents=True,exist_ok=True)
    if a.seal and not a.go: raise SystemExit('--seal requires --go')
    entries,dirs,excluded=collect(packet)
    if a.manifest_only:
        status='planning_manifest_not_sealed'
    elif not a.seal:
        status='planning_only_no_archive'
    else:
        status='sealed'
    manifest={'schema':SCHEMA,'status':status,'generated_utc':datetime.now(timezone.utc).isoformat(),'packet_root':str(packet),'entry_count':len(entries),'logical_bytes':sum(x['bytes'] for x in entries),'entries':entries,'directories':dirs,'excluded':excluded,'symlinks':[],'dependencies':{'0950_archive':{'path':'/home/zhuhe/litchi-performance-0950-archive/archive-0950-v1.tar.zst','sha256':'cee0b5d2c9e878ca577994d4a4f8e56032c1d119fe6e27ac3b3d8a8e51daa363','copied':False},'0944_archive':{'path':'/home/zhuhe/litchi-performance-0944-archive/litchi-performance-0944-v1.tar.zst','sha256':'3f43604846b87843bbc71fcd4c8f075d57219022a1c737c472973186e85fb592','copied':False}},'cleanup_requires_explicit_go':True}
    mpath=root/'manifest-v1.json'; mpath.write_text(json.dumps(manifest,indent=2,sort_keys=True)+'\n')
    if not a.seal:
        print(json.dumps({'status':status,'manifest':str(mpath),'entries':len(entries),'bytes':manifest['logical_bytes']})); return
    for name in ('archive-0951-v1.tar','archive-0951-v1.tar.zst','selection-v1.list','archive-receipt-v1.json'):
        if (root/name).exists(): raise RuntimeError(f'refusing existing artifact {name}')
    selection=root/'selection-v1.list'; selection.write_bytes(b''.join(x['path'].encode()+b'\0' for x in dirs+entries))
    tar=root/'archive-0951-v1.tar'; zst=root/'archive-0951-v1.tar.zst'; tmp=root/'.archive-0951-v1.tar.zst.tmp'
    subprocess.run(['/usr/bin/tar','--create','--no-recursion','--sort=name','--mtime=@0','--owner=0','--group=0','--numeric-owner','--format=pax','--file',str(tar),'--null','--verbatim-files-from','--files-from',str(selection)],cwd=packet,check=True)
    subprocess.run(['/usr/bin/zstd','-3','--no-progress','--force',str(tar),'-o',str(tmp)],check=True); tmp.replace(zst)
    td=digest(tar); tar.unlink(); zd=digest(zst)
    receipt={'schema':'litchi.performance.0951.external-archive-receipt.v1','status':'sealed','manifest':digest(mpath),'selection':digest(selection),'compressed_archive':zd,'uncompressed_tar':{'path':str(tar),'bytes':td['bytes'],'sha256':td['sha256'],'deleted_after_compression':True},'helper':digest(Path(__file__).resolve()),'packet_census':{'entries':len(entries),'logical_bytes':manifest['logical_bytes'],'directories':len(dirs),'excluded':excluded,'symlinks':[]},'dependencies':manifest['dependencies'],'regeneration':f'python3 -B {Path(__file__).resolve()} --packet {packet} --archive-root {root} --seal --go'}
    rpath=root/'archive-receipt-v1.json'; rpath.write_text(json.dumps(receipt,indent=2,sort_keys=True)+'\n')
    print(json.dumps({'status':'sealed','manifest':str(mpath),'archive':str(zst),'entries':len(entries),'logical_bytes':manifest['logical_bytes'],'archive_bytes':zd['bytes'],'archive_sha256':zd['sha256']}))
if __name__=='__main__': main()
