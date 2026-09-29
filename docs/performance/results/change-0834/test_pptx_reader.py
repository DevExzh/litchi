"""Reject forged aligned replay claims while retaining raw metadata overlap."""
from copy import deepcopy
from pathlib import Path
import hashlib,json,sys
import reader

P=Path(__file__).resolve().parent
def main():
 path=P/'qualification-v2-05.json';base=json.loads(path.read_text())
 binary=json.loads((P/'build-repaired-v2.json').read_text())['binary']
 def validate(value):reader.validate_report(value,reader.CASES[5],reader.STATES,1,0,binary)
 def cold(value):return next(s for s in value['filesystem_evidence'][0]['samples'] if s['cache_state']=='cold-verified')['pptx_source_replay']
 def proof(value):return cold(value)['aligned_eocd_tail_probe']
 def wrong_semantics(value):
  for s in value['filesystem_evidence'][0]['samples']:s['pptx_source_replay']['semantic_sha256']='0'*64
 validate(base)
 mutations={
  'missing-aligned-proof':lambda v:cold(v).pop('aligned_eocd_tail_probe'),
  'wrong-base-hash':lambda v:proof(v)['alignment_proof'].update(base_sha256='0'*64),
  'wrong-padding':lambda v:proof(v)['alignment_proof'].update(padding_bytes=0),
  'wrong-eocd':lambda v:proof(v)['alignment_proof'].update(eocd_offset=0),
  'wrong-boundary':lambda v:proof(v).update(open_read_count=1),
  'missing-tail':lambda v:proof(v)['raw_reads'].pop(1),
  'duplicate-tail':lambda v:proof(v)['raw_reads'].insert(1,deepcopy(proof(v)['raw_reads'][1])),
  'hidden-raw-overlap':lambda v:cold(v).update(unselected_slide_payload_read_bytes=0),
  'overstated-return':lambda v:proof(v)['raw_reads'][1].update(returned_length=65537),
  'wrong-selected-range':lambda v:proof(v)['payload_ranges']['selected_slide'].update(start=0),
  'wrong-semantics-in-both-states':wrong_semantics,
 }
 results=[]
 for name,mutate in mutations.items():
  value=deepcopy(base);mutate(value)
  try:validate(value)
  except reader.QualificationError:results.append(dict(name=name,rejected=True))
  else:raise AssertionError(f'accepted mutation: {name}')
 result=dict(status='pass',report_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),mutations=results)
 dest=P/'pptx-reader-tests.json'
 if '--check' in sys.argv:assert json.loads(dest.read_text())==result
 else:
  with dest.open('x') as f:json.dump(result,f,indent=2);f.write('\n')
 print('PPTX reader mutations PASS:',len(results))
if __name__=='__main__':main()
