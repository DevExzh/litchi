"""Compare the binary-generated catalog with the independently frozen literal inputs."""
import hashlib
import custody as c

def check():
 expected=c.read(c.P/'fixtures.json');actual=c.read(c.P/'cases.json')
 assert len(expected)==len(actual)==33
 for e,a in zip(expected,actual):
  assert all(e[k]==a[k] for k in ['id','category','attribute_count'])
  s=a['source'];raw=s['value'].encode() if s['encoding']=='utf8' else bytes.fromhex(s['value'])
  assert raw==e['input'].encode() and len(raw)==s['bytes']==e['bytes'] and hashlib.sha256(raw).hexdigest()==e['sha256']
if __name__=='__main__':check()
