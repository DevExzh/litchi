"""Small independent ZIP-only preservation cross-check of all qualifications."""
import driver as d
import zipfile,io,hashlib

def verify():
 results=[]
 for row in d.read(d.P/'plan.json')['qualification']:
  folder=d.P/'runs'/f"qualification-00-{row['parts']:04}-{row['mode']}"
  source=(folder/'source.zip').read_bytes();output=(folder/'artifact.zip').read_bytes();n=row['parts']
  expected={f'custom/p{i:05}.bin' for i in range(n)}|{'[Content_Types].xml','_rels/.rels'}
  with zipfile.ZipFile(io.BytesIO(source)) as a,zipfile.ZipFile(io.BytesIO(output)) as b:
   assert set(a.namelist())==set(b.namelist())==expected and len(a.infolist())==len(b.infolist())==n+2
   for i in range(n):
    name=f'custom/p{i:05}.bin';payload=bytes((i*17+j*13)%251 for j in range(256));changed=bytes([payload[0]^1])+payload[1:] if i==n-1 else payload
    assert a.read(name)==payload and b.read(name)==changed
   for name in ['[Content_Types].xml','_rels/.rels']:assert a.read(name)==b.read(name)
   def spans(z,raw):
    infos=sorted(z.infolist(),key=lambda x:x.header_offset)
    return {x.filename:raw[x.header_offset:(infos[i+1].header_offset if i+1<len(infos) else z.start_dir)] for i,x in enumerate(infos)}
   sa,sb=spans(a,source),spans(b,output);target=f'custom/p{n-1:05}.bin';unchanged=sorted(expected-{target})
   fields=['date_time','compress_type','flag_bits','CRC','compress_size','file_size','extra','comment','create_system','create_version','extract_version','reserved','volume','internal_attr','external_attr']
   for name in unchanged:
    assert sa[name]==sb[name]
    assert all(getattr(a.getinfo(name),k)==getattr(b.getinfo(name),k) for k in fields)
   assert a.read(target)!=b.read(target)
  results.append(dict(parts=n,mode=row['mode'],members=n+2,unchanged_spans=n+1,source=d.desc(folder/'source.zip'),output=d.desc(folder/'artifact.zip')))
 return dict(status='pass',rows=results,checker=d.desc(d.P/'archive_check.py'))
if __name__=='__main__':
 value=verify();p=d.P/'archive-check.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Independent ZIP cross-check PASS',len(value['rows']),'source/output pairs')
