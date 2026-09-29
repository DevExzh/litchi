"""Mutation checks for per-route cold OPC framing evidence."""
from copy import deepcopy
from pathlib import Path
import hashlib,json,sys
import reader

P=Path(__file__).resolve().parent
def main():
 path=P/'qualification-v3-03.json';base=json.loads(path.read_text())
 binary=json.loads((P/'build-repaired-v3.json').read_text())['binary']
 def validate(value):reader.validate_report(value,reader.CASES[3],reader.STATES,1,0,binary)
 def evidence(value):return value['filesystem_evidence'][0]
 def alignment(value):return evidence(value)['opc_cold_alignment_proof']
 def output(value):return evidence(value)['opc_cold_output_proof']
 validate(base)
 mutations={
  'missing-output-proof':lambda v:evidence(v).pop('opc_cold_output_proof'),
  'wrong-base-hash':lambda v:alignment(v).update(base_sha256='0'*64),
  'wrong-padding':lambda v:alignment(v).update(padding_bytes=0),
  'wrong-eocd-position':lambda v:alignment(v).update(eocd_offset=0),
  'wrong-route':lambda v:output(v).update(route='eager'),
  'wrong-canonical-hash':lambda v:output(v).update(canonical_sha256='0'*64),
  'wrong-output-hash':lambda v:output(v).update(output_sha256='0'*64),
  'wrong-output-length':lambda v:output(v).update(output_bytes=0),
  'wrong-comment-length':lambda v:output(v).update(output_comment_bytes=0),
  'missing-unchanged-member':lambda v:output(v).update(unchanged_member_count=4),
 }
 rows=[]
 for name,mutate in mutations.items():
  value=deepcopy(base);mutate(value)
  try:validate(value)
  except reader.QualificationError:rows.append(dict(name=name,rejected=True))
  else:raise AssertionError(f'accepted mutation: {name}')
 result=dict(status='pass',report_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),mutations=rows)
 dest=P/'opc-reader-tests.json'
 if '--check' in sys.argv:assert json.loads(dest.read_text())==result
 else:
  with dest.open('x') as f:json.dump(result,f,indent=2);f.write('\n')
 print('OPC reader mutations PASS:',len(rows))
if __name__=='__main__':main()
