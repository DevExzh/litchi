#!/usr/bin/env python3
from __future__ import annotations
import hashlib,json,os,shutil,stat,sys
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0948').resolve()
AR=Path('/home/zhuhe/litchi-performance-0948-archive')
SEL=AR/'target-cleanup-selection-v3.json'
START=AR/'target-cleanup-start-v1.json'
OUT=AR/'target-cleanup-receipt-v1.json'
ALLOWED={ROOT/'target-candidate-native-v3/release/litchi-perf-baseline',ROOT/'target-candidate-allocator-v3/release/litchi-perf-baseline-alloc'}

def sha(p):
 h=hashlib.sha256();
 with p.open('rb') as f:
  while b:=f.read(1024*1024): h.update(b)
 return h.hexdigest()
def desc(p):
 st=p.stat(); return {'path':str(p.relative_to(ROOT)),'bytes':st.st_size,'mode_bits':stat.S_IMODE(st.st_mode),'sha256':sha(p)}
def fail(msg): raise RuntimeError(msg)
def validate_path(p):
 try: rel=p.relative_to(ROOT)
 except ValueError: fail(f'outside root: {p}')
 if p.is_symlink(): fail(f'symlink selected: {rel}')
 if not (rel.parts[0].startswith('target-candidate-') or rel.parts[0]=='quality'):
  fail(f'not target cleanup namespace: {rel}')
 if p.is_dir() and any(p == a or p in a.parents for a in ALLOWED):
  fail(f'selected root contains preserved executable: {rel}')

def main():
 selection=json.loads(SEL.read_text())
 rows=selection['disposable_roots']
 if selection['status']!='inventory_only_not_deleted': fail('selection is not inventory-only')
 if selection['root']!=str(ROOT): fail('selection root mismatch')
 # All selected roots must still have exactly the recorded inventory before deletion.
 pre=[]
 for row in rows:
  p=ROOT/row['path']; validate_path(p)
  if not p.exists(): fail(f'missing selected root: {row["path"]}')
  if p.is_file(): files=1; dirs=0; bytes_=p.stat().st_size
  else:
   files=dirs=bytes_=0
   for dp,dn,fn in os.walk(p,followlinks=False):
    dirs+=len(dn)
    for n in dn:
     if (Path(dp)/n).is_symlink(): fail(f'symlink directory: {Path(dp)/n}')
    for n in fn:
     q=Path(dp)/n
     if q.is_symlink(): fail(f'symlink file: {q}')
     if q.is_file(): files+=1; bytes_+=q.stat().st_size
  if (files,dirs,bytes_) != (row['files'],row['dirs'],row['bytes']): fail(f'pre-delete inventory changed: {row["path"]}')
  pre.append({'path':row['path'],'files':files,'dirs':dirs,'bytes':bytes_})
 start={'schema':'litchi.performance.0948.target-cleanup-start.v1','status':'authorized_started','root':str(ROOT),'selection':str(SEL),'selection_sha256':sha(SEL),'preserved_before':{str(a.relative_to(ROOT)):desc(a) for a in sorted(ALLOWED)},'roots':pre}
 START.write_text(json.dumps(start,indent=2,sort_keys=True)+'\n')
 removed=[]
 try:
  for row in rows:
   p=ROOT/row['path']; validate_path(p)
   if p.is_dir(): shutil.rmtree(p)
   elif p.is_file(): p.unlink()
   else: fail(f'selected path vanished: {row["path"]}')
   removed.append(row['path']); print('removed',len(removed),row['path'],flush=True)
  remaining=[row['path'] for row in rows if (ROOT/row['path']).exists() or (ROOT/row['path']).is_symlink()]
  preserved={str(a.relative_to(ROOT)):desc(a) for a in sorted(ALLOWED)}
  checks={'all_selected_absent':not remaining,'remaining_selected':remaining,'preserved_after':preserved,'preserved_unchanged':preserved==start['preserved_before']}
  status='pass' if checks['all_selected_absent'] and checks['preserved_unchanged'] else 'fail'
  result={'schema':'litchi.performance.0948.target-cleanup-receipt.v1','status':status,'root':str(ROOT),'selection':str(SEL),'selection_sha256':sha(SEL),'selection_roots':len(rows),'removed_roots':removed,'removed_file_count':sum(x['files'] for x in rows),'removed_bytes':sum(x['bytes'] for x in rows),'checks':checks,'preserved_paths':sorted(checks['preserved_after'])}
 except Exception as e:
  result={'schema':'litchi.performance.0948.target-cleanup-receipt.v1','status':'failed_partial','root':str(ROOT),'selection':str(SEL),'selection_sha256':sha(SEL),'removed_roots':removed,'error':repr(e)}
 OUT.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
 print(json.dumps(result,indent=2,sort_keys=True))
 return 0 if result['status']=='pass' else 1
if __name__=='__main__': raise SystemExit(main())
