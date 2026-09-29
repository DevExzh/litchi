"""Add changed-target stable metadata checks to the retained ZIP cross-check."""
import archive_check as original
import driver as d
import zipfile

def verify():
 prior=original.verify()
 assert prior==d.read(d.P/'archive-check.json')
 fields=['date_time','compress_type','flag_bits','file_size','extra','comment','create_system','create_version','extract_version','reserved','volume','internal_attr','external_attr']
 for row in prior['rows']:
  with zipfile.ZipFile(row['source']['path']) as a,zipfile.ZipFile(row['output']['path']) as b:
   name=f"custom/p{row['parts']-1:05}.bin"
   assert all(getattr(a.getinfo(name),k)==getattr(b.getinfo(name),k) for k in fields)
 return dict(status='pass',original=d.desc(d.P/'archive-check.json'),checker=d.desc(d.P/'archive_check_v2.py'),pairs=len(prior['rows']),target_stable_fields=fields)
if __name__=='__main__':
 value=verify();p=d.P/'archive-check-v2.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Target stable metadata PASS',value['pairs'],'pairs')
