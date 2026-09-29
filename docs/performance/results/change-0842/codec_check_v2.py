"""Add explicit Deflate-route and same-leg build checks to codec qualification."""
import driver as d
import qualify as q
import codec_check as original

def verify():
 value=original.verify(['before','after'])
 for leg in ['before','after']:
  receipt=d.read(d.P/f'codecs-{leg}.json')
  assert receipt['build']==d.desc(d.P/f'build-{leg}.json')
  build=d.read(q.bind(receipt['build']));frozen=d.read(q.bind(build['freeze']))
  assert frozen['source']==d.read(d.P/f'source-{leg}.json')['source']
  for row in value['results'][leg]:
   for member in row['manifest']['members']:
    assert member['compression']=='deflated'
    assert member['central']['compress_type']==member['local']['method']==8
 value.update(explicit_deflate_methods=True,same_leg_builds=True,additional_checker=d.desc(d.P/'codec_check_v2.py'))
 return value
if __name__=='__main__':
 value=verify();out=d.P/'codecs-validation-v2.json'
 if out.exists():assert d.read(out)==value
 else:d.write(out,value)
 print('Explicit Deflate and same-leg codec qualification PASS;18archives')
