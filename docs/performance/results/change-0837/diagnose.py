"""Replay the durable inverse byte discrepancy from retained assertion artifacts."""
import hashlib,json,struct,zipfile
from pathlib import Path
P=Path(__file__).resolve().parent
record=json.loads((P/'inverse-failure-analysis.json').read_text())
paths={'left':P/'diagnostics/durable-restored.zip','right':P/'diagnostics/fresh-source.zip'}
for side,key in [('left','restored'),('right','source')]:
 raw=paths[side].read_bytes();assert len(raw)==record[key]['bytes'];assert hashlib.sha256(raw).hexdigest()==record[key]['sha256']
archives={side:zipfile.ZipFile(path) for side,path in paths.items()};assert archives['left'].namelist()==archives['right'].namelist()
facts=[]
for name in archives['left'].namelist():
 payloads={side:archive.read(name) for side,archive in archives.items()};assert payloads['left']==payloads['right']
 compressed={}
 for side,archive in archives.items():
  member=archive.getinfo(name);raw=paths[side].read_bytes();offset=member.header_offset
  namelen,extra=struct.unpack_from('<HH',raw,offset+26);start=offset+30+namelen+extra;span=raw[start:start+member.compress_size]
  compressed[side]=dict(bytes=len(span),sha256=hashlib.sha256(span).hexdigest(),method=member.compress_type)
 facts.append(dict(name=name,payload_bytes=len(payloads['left']),payload_sha256=hashlib.sha256(payloads['left']).hexdigest(),compressed=compressed,compressed_equal=compressed['left']==compressed['right']))
assert facts==record['members'] and record['all_member_payloads_identical'] is True and record['archive_bytes_equal'] is False
assert [r['name'] for r in facts if not r['compressed_equal']]==['word/document.xml']
print('diagnostic replay PASS: all logical bytes equal; word/document.xml compressed span differs')
