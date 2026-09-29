"""Replay every retained qualification stage without promoting failures."""
from pathlib import Path
import hashlib,json,sys
import reader

P=Path(__file__).resolve().parent
def read(path):return json.loads(path.read_text())
def compute():
 diagnostic=P/'diagnostic-pptx-warm.json'
 reader.validate_report(diagnostic,reader.CASES[5],['warm'],1,0,read(P/'build-diagnostic.json')['binary'])
 failed=reader.validate_manifest(P/'qualification-v2.json')
 final=reader.validate_manifest(P/'qualification-v3.json')
 assert failed['status']=='failed' and not failed['qualification_valid']
 assert len(failed['reports'])==4 and failed['validated_sample_count']==8 and len(failed['failed_rows'])==3
 assert final['status']=='pass' and final['qualification_valid']
 assert len(final['reports'])==7 and final['validated_sample_count']==16
 return dict(status='pass',reports=12,samples=25,formal_reports=0,formal_samples=0,diagnostic_warm_sha256=hashlib.sha256(diagnostic.read_bytes()).hexdigest(),failed_v2=failed,final_v3=final)
def main():
 value=compute();path=P/'reader-replay.json'
 if '--check' in sys.argv:assert read(path)==value
 else:
  with path.open('x') as f:json.dump(value,f,indent=2,sort_keys=True);f.write('\n')
 print('reader replay PASS: 12 reports / 25 diagnostic or qualification samples; zero formal samples')
if __name__=='__main__':main()
