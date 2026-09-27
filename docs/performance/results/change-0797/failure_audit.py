"""Retain and verify the terminal pre-capture test compilation failure."""
from pathlib import Path
import custody as c
p=c.P;q=p/'quality-failed-0';w=c.read(q/'relocation.json');rows=c.read(q/'receipts.json')
assert len(rows)==5 and [r['exit_code']==0 for r in rows]==[True,True,True,True,False]
last=0
for r in rows:
 assert last<=r['started']<=r['ended'];last=r['ended']
 a=r['log'];log=q/Path(a['path']).name;assert c.sha(log)==a['sha256'] and log.stat().st_size==a['bytes']
assert 'error[E0631]' in (q/'after-1.log').read_text()
assert last<=c.read(p/'quality/complete.json')['rows'][0]['started']
for name,h in c.read(q/'archive-inputs.json').items():
 if name.startswith(('before/','after/')):
  leg,file=name.split('/',1)
  if file=='litchi-opc-xml_attributes-tests.rs':f=p/'test-src-failed-0'/leg/'crates/litchi-opc/src/xml_attributes/tests.rs'
  else:f=p/'test-src-failed-0'/leg/'crates'/file.removesuffix('-xml_attributes.rs')/'src/xml_attributes.rs'
 else:f=p/'candidate'/name
 assert c.sha(f)==h,name
failed=(p/'test-src-failed-0/after/crates/litchi-opc/src/xml_attributes/tests.rs').read_text()
fixed=(p/'candidate/after/litchi-opc-xml_attributes-tests.rs').read_text()
assert failed.replace('any(Result::is_err)','any(|item| item.is_err())')==fixed
assert w['production_changed'] is False
print('Retained test-only compilation failure custody PASS')
