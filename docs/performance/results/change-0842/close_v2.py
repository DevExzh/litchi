"""Additive closure: retain original audit and require separate tiny-tail evidence."""
import driver as d
import qualify as q
import tail_analyze as t
import runpy,sys
for item in d.read(d.P/'tail-admission.json')['files']:q.bind(item)
v=t.analyze();assert v==d.read(d.P/'tail-analysis.json') and v['reports']==24 and v['samples']==24000
check=d.read(d.P/'tail-crosscheck.json');assert check['status']=='pass' and check['metric_rows']==8 and check['paired_intervals']==8
for key in ['analysis','checker']:q.bind(check[key])
decision=d.read(d.P/'decision.json');assert decision['tail_analysis']==d.desc(d.P/'tail-analysis.json')
assert decision['tail_confirmation_flags']==len(v['regression_flags'])
if decision['decision']=='adopt':assert not v['regression_flags']
runpy.run_path(str(d.P/'close.py'),run_name='__main__')
