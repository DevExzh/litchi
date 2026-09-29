"""Mutation tests for the independent diagnostic reader."""
from copy import deepcopy
from pathlib import Path
import json,sys
import audit

P=Path(__file__).resolve().parent

def main():
 original=audit.parse_nested_diagnostic
 log=P/'commands/diagnostic-pptx-cold-verified/output.log'
 base=original(log)
 audit.validate_raw_diagnostic(log)
 mutations={
  'wrong-child-mode':lambda d:d.update(child_mode='cold-verified'),
  'wrong-source-hash':lambda d:d.update(source_sha256='0'*64),
  'shifted-open-boundary':lambda d:d.update(open_read_count=1),
  'missing-tail':lambda d:d['raw_reads'].pop(1),
  'duplicate-tail':lambda d:d['raw_reads'].insert(1,deepcopy(d['raw_reads'][1])),
  'short-tail':lambda d:d['raw_reads'][1].update(returned_length=65535),
  'oversized-return':lambda d:d['raw_reads'][1].update(returned_length=65537),
  'counter-mismatch':lambda d:d['counters'].update(read_bytes=d['counters']['read_bytes']+1),
  'coverage-mismatch':lambda d:d['coverage'].update(unselected_slide_payload_covered_bytes=0),
  'selected-range-mismatch':lambda d:d['payload_ranges']['selected_slide'].update(start=0),
 }
 results=[]
 try:
  for name,mutate in mutations.items():
   candidate=deepcopy(base);mutate(candidate)
   audit.parse_nested_diagnostic=lambda unused,candidate=candidate:deepcopy(candidate)
   try:audit.validate_raw_diagnostic(log)
   except (audit.AuditError,AssertionError,ValueError,TypeError,KeyError,IndexError):results.append(dict(name=name,rejected=True))
   else:raise AssertionError(f'accepted mutation: {name}')
 finally:audit.parse_nested_diagnostic=original
 value=dict(status='pass',diagnostic_mutations=results)
 path=P/'reader-tests.json'
 if '--check' in sys.argv:assert json.loads(path.read_text())==value
 else:
  with path.open('x') as f:json.dump(value,f,indent=2);f.write('\n')
 print('reader mutation tests PASS:',len(results))

if __name__=='__main__':main()
