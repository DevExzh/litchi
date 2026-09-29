"""Add independent confirmation replay without altering the frozen main closure."""
import driver as d
import qualify as q
import tail,tail_crosscheck,tail_tables
import runpy
value=tail.analyze();assert value==d.read(d.P/'tail-analysis.json')
assert value['reports']==72 and value['samples']==72000
plan=d.read(d.P/'tail-plan.json');original=d.read(d.P/'analysis.json');cases=[]
for x in original['tail_review']:
 c={k:x[k] for k in ['shape','parts','mode']}
 if c not in cases:cases.append(c)
assert len(original['tail_review'])==6 and plan['cases']==cases
for case in cases:
 first=[next(x for x in plan['rows'] if x['block']==b and all(x[k]==v for k,v in case.items()))['leg'] for b in range(6)]
 assert first.count('before')==first.count('after')==3
assert all(x['samples']==1000 and x['warmup']==3 for x in plan['rows'])
capture=d.read(d.P/'capture-tail.json');assert capture['status']=='pass' and capture['plan']==d.desc(d.P/'tail-plan.json')
assert capture['receipts']==[d.desc(tail.folder(row)/'receipt.json') for row in plan['rows']]
assert tail_crosscheck.verify()==d.read(d.P/'tail-crosscheck.json')
assert tail_tables.render()==(d.P/'tail-tables.md').read_text()
decision=d.read(d.P/'decision.json');assert decision['tail_confirmation']==d.desc(d.P/'tail-analysis.json')
assert decision['tail_confirmation_flags']==len(value['regression_flags']) and decision['tail_confirmation_points']==len(value['tail_points']) and decision['tail_confirmation_spreads']==len(value['spread_flags'])
if decision['decision']=='adopt':assert not value['regression_flags'] and not value['spread_flags'] and decision['tail_confirmation_reviewed']
runpy.run_path(str(d.P/'close.py'),run_name='__main__')
