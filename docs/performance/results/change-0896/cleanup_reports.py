#!/usr/bin/env python3
"""Remove only hash-verified raw reports after reviewed gzip roundtrips."""
from pathlib import Path
import gzip,hashlib,json
P=Path(__file__).resolve().parent
def digest(f):
 h=hashlib.sha256();size=0
 with f as stream:
  for b in iter(lambda:stream.read(1024*1024),b''):h.update(b);size+=len(b)
 return size,h.hexdigest()
def main():
 v=json.loads((P/'corpus-comparison.json').read_text());review=json.loads((P/'corpus-review.json').read_text());assert v['status']==review['status']=='pass'
 rows=[]
 for value in v['compressed_reports'].values():
  raw=value['source'];compressed=value['compressed'];f=Path(raw['path']);g=Path(compressed['path']);assert f.parent==g.parent==P and not f.is_symlink() and not g.is_symlink()
  assert digest(f.open('rb'))==(raw['bytes'],raw['sha256'])
  assert digest(g.open('rb'))==(compressed['bytes'],compressed['sha256'])
  assert digest(gzip.open(g,'rb'))==(raw['bytes'],raw['sha256'])
  rows.append({'raw':raw,'compressed':compressed,'roundtrip':True})
 old=v['diagnostic_repeat']['report'];assert v['diagnostic_repeat']['formal_baseline'] is False
 assert digest(Path(old['path']).open('rb'))==(old['bytes'],old['sha256'])
 assert (old['bytes'],old['sha256'])==(rows[0]['raw']['bytes'],rows[0]['raw']['sha256'])
 rows.append({'raw':old,'compressed':rows[0]['compressed'],'roundtrip':True,'deduplicated_equal_bytes':True,'formal_baseline':False})
 for row in rows:Path(row['raw']['path']).unlink();row['raw_absent']=not Path(row['raw']['path']).exists()
 with (P/'reports-cleanup.json').open('x') as f:json.dump({'status':'pass','rows':rows,'reclaimed_bytes':sum(r['raw']['bytes'] for r in rows),'helper_sha256':digest(Path(__file__).open('rb'))[1]},f,indent=2);f.write('\n')
 print('Raw report cleanup PASS',sum(r['raw']['bytes'] for r in rows))
if __name__=='__main__':main()
