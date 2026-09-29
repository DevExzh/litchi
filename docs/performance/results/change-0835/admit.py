"""Bind completed independent qualification or capture validation exactly once."""
import sys,time
import driver as d

def main(mode):
 assert mode in ['qualification','capture']
 qualification = mode=='qualification'
 validation='qualification-validation-v2.json' if qualification else 'capture-validation.json'
 value=d.read(d.P/validation)
 assert value['status']=='pass'
 assert value['qualification_valid' if qualification else 'capture_valid'] is True
 assert value['validated_sample_count']==(16 if qualification else 2160)
 assert d.read(d.P/'reader-tests.json')['status']=='pass'
 names=['origin.json','host.json','freeze-baseline.json','build-baseline.json','quality-reuse.json','measurement-plan.json','reader.py','reader-tests.py','reader-tests.json','qualification.json',validation,'capture.py','analyze.py','audit.py','preflight.py','preflight.json','reader-correction.json','measurement-plan-initial.json']
 if not qualification:names+=['capture.json','capture-started.json','admission.json']
 d.write(d.P/('admission.json' if qualification else 'capture-admission.json'),dict(status='pass',created_unix=time.time(),inputs={n:d.desc(d.P/n) for n in names},validated_sample_count=value['validated_sample_count'],performance_claim='descriptive route/cache baseline only'))
 print(mode,'admission PASS')
if __name__=='__main__':main(sys.argv[1])
