"""Replay the v2 suite with an explicit binding to its actual imported reader."""
import driver as d
import reader_tests_v2 as suite
reader=d.Path(suite.r.__file__).resolve();assert reader==d.P/'readers_v2.py'
before=d.desc(reader);value=suite.run();assert d.desc(reader)==before
assert value['status']=='pass' and value['mutation_count']==21
value['inherited_receipt_reader_label']=value.pop('reader')
value['reader']=before
value['test_implementation']=d.desc(d.P/'reader_tests_v2.py')
value['test_runner']=d.desc(d.P/'reader_tests_final.py')
value['note']='The frozen v2 suite imports readers_v2 but inherited its old readers.py receipt label. This additive replay binds the actual imported module; code, corpus, metrics and workloads are unchanged.'
out=d.P/'reader-tests-final.json'
if out.exists():assert d.read(out)==value
else:d.write(out,value)
print('Final reader replay PASS;21checks;actual imported reader bound')
