#!/usr/bin/env python3
"""Offline controls use existing reports with synthetic zero-warmup metadata."""
import copy,importlib.util,json,subprocess,sys,unittest
from pathlib import Path
ROOT=Path('/home/zhuhe/litchi-performance-0962')
spec=importlib.util.spec_from_file_location('capture',ROOT/'tools/capture_fp_v1.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
class Controls(unittest.TestCase):
 def test_reports_and_corruptions(self):
  cases=json.loads((ROOT/'fp-plan-v1.json').read_text())['cases']
  for c in cases:
   v=json.loads((ROOT/'cpu-profile-native-v1'/c['case_id']/'output/report.json').read_text())
   v['warmup_iterations']=v['case']['warmup_iterations']=0
   n=v['measured_samples'];m.validate(v,c,n)
   for field,value in [('output_sha256','bad'),('semantic_sha256','bad'),('cpu_tasks_used',999),('fixture_inverse_restored_exact_source',False),('sink_largest_write',65537),('total_ns',1)]:
    bad=copy.deepcopy(v);bad['case']['samples'][0][field]=value
    with self.assertRaises(AssertionError,msg=(c['case_id'],field)):m.validate(bad,c,n)
   bad=copy.deepcopy(v);bad['case']['samples'][0]['resources_after_drop']['memory']=1
   with self.assertRaises(AssertionError):m.validate(bad,c,n)
   bad=copy.deepcopy(v);bad['case']['source_sha256']='bad'
   with self.assertRaises(AssertionError):m.validate(bad,c,n)
 def test_optimized_python_refused(self):
  p=subprocess.run([sys.executable,'-O',str(ROOT/'tools/capture_fp_v1.py'),'--help'],capture_output=True,text=True)
  self.assertNotEqual(p.returncode,0);self.assertIn('optimized Python is not supported',p.stderr)
if __name__=='__main__':unittest.main()
