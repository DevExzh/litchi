"""Independent exact codec-route differential; no timings inferred."""
import driver as d
import reader as r
import qualify as q
import hashlib,io,zipfile,sys
ROUTES=['borrowed','owned-write-1','owned-write-7','owned-write-32768','sized','reader-stream','reader-alias','reusable','preserved-copy']
def expected_payloads():
 size=32768+257;pattern=b'litchi-0842-codec-pattern-0123456789abcdef';state=0x9e3779b9;data=bytearray()
 for _ in range(size):
  state=(state^(state<<13))&0xffffffff;state=(state^(state>>17))&0xffffffff;state=(state^(state<<5))&0xffffffff;data.append(state>>24)
 return {'codec/empty.bin':b'','codec/tiny.bin':b'soapberry-0842-tiny','codec/pattern.bin':bytes(pattern[i%len(pattern)] for i in range(size)),'codec/random.bin':bytes(data)}
def verify(legs):
 expected=expected_payloads();results={}
 digest=hashlib.sha256()
 for n,b in expected.items():
  digest.update(len(n).to_bytes(8,'little'));digest.update(n.encode());digest.update(len(b).to_bytes(8,'little'));digest.update(b)
 logical=dict(bytes=sum(map(len,expected.values())),sha256=digest.hexdigest())
 for leg in legs:
  receipt=d.read(d.P/f'codecs-{leg}.json');build=d.read(q.bind(receipt['build']));command=d.read(q.bind(receipt['command']));q.bind(command['log'])
  assert command['exit_code']==0 and command['argv']==['taskset','-c','12',build['codec']['path'],str(d.P/'codecs'/leg)]
  for x in receipt['artifacts']:q.bind(x)
  folder=d.P/'codecs'/leg;summary=d.read(folder/'summary.json')
  assert summary['schema']=='litchi.compression-codec-qualification-0842.v1' and summary['tool']=='compression-probe-0842-codecs'
  assert [x['route'] for x in summary['routes']]==ROUTES and summary['expected_logical']==logical
  assert {x.name for x in folder.iterdir()}=={x+'.zip' for x in ROUTES}|{'summary.json'}
  assert [x['name'] for x in summary['payloads']]==list(expected)
  for x in summary['payloads']:
   payload=expected[x['name']];assert x['bytes']==len(payload) and x['payload_sha256']==hashlib.sha256(payload).hexdigest()
  results[leg]=[]
  for row in summary['routes']:
   blob=(folder/(row['route']+'.zip')).read_bytes();manifest=r.archive_manifest(blob)
   assert row['archive']==manifest['archive'] and row['logical']==row['expected_logical']==logical
   with zipfile.ZipFile(io.BytesIO(blob)) as z:
    assert z.namelist()==list(expected)
    assert all(z.read(n)==b for n,b in expected.items())
   assert len(row['members'])==4
   for x,(name,payload) in zip(row['members'],expected.items()):
    assert x==dict(name=name,bytes=len(payload),expected_payload_sha256=hashlib.sha256(payload).hexdigest(),decoded_payload_sha256=hashlib.sha256(payload).hexdigest())
   results[leg].append(dict(route=row['route'],manifest=manifest))
  # These named controls intentionally exercise the same repeated writer API.
  assert (folder/'borrowed.zip').read_bytes()==(folder/'reusable.zip').read_bytes()==(folder/'preserved-copy.zip').read_bytes()
 if len(legs)==2:
  assert results['before']==results['after']
  assert all((d.P/'codecs/before'/(route+'.zip')).read_bytes()==(d.P/'codecs/after'/(route+'.zip')).read_bytes() for route in ROUTES)
 return dict(status='pass',processes=len(legs),archives=9*len(legs),members=36*len(legs),results=results,reader=d.desc(d.P/'reader.py'),checker=d.desc(d.P/'codec_check.py'),limitation='Nine named artifacts include a duplicate borrowed/reusable control; this is not nine distinct algorithms.')
if __name__=='__main__':
 legs=['before'] if sys.argv[1:]==['before'] else ['before','after'];value=verify(legs)
 out=d.P/('codecs-validation-before.json' if len(legs)==1 else 'codecs-validation.json')
 if out.exists():assert d.read(out)==value
 else:d.write(out,value)
 print('codec byte qualification PASS',value['archives'],'archives')
