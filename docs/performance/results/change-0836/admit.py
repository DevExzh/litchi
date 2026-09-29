"""Freeze native admission only after build, symbol, source and report checks."""
import time
import driver as d
import reader

def main():
 for s in ['baseline','fp']:d.check(s);reader._prepare_binding(s)
 assert d.read(d.P/'symbols.json')['status']=='pass'
 assert d.read(d.P/'reader-preflight-tests.json')['status']=='pass'
 q=reader.validate_manifest(d.P/'qualification.json');assert q['status']=='pass' and q['validated_sample_count']==4
 names=['origin.json','plan.json','quality-reuse.json','host.json','tools.json','driver.py','runner.py','capture.py','reader.py','reader-tests.py','reader-preflight-tests.json','qualification.json','qualification-validation.json','symbols.py','symbols.json','freeze-baseline.json','freeze-fp.json','build-baseline.json','build-fp.json']
 d.write(d.P/'admission.json',dict(status='pass',created_unix=time.time(),inputs={n:d.desc(d.P/n) for n in names},scope='Native controls admitted; CPU/trace artifact readers and independent final audit remain required before attribution.'))
 print('native admission PASS')
if __name__=='__main__':main()
