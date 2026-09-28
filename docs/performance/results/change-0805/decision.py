"""Frozen micro-preflight advancement; never production adoption."""
import sys
import custody as c

p=c.P
plan=c.read(p/'plan.json')
assert plan['schema']=='litchi.performance.0805.v1'
assert plan['analysis']['bootstrap_seed']==805080
policy=plan['preflight_decision']
assert policy['advance_requires_semantics'] is True
assert policy['benefit_cases']==['distinct-1','distinct-2']
assert policy['benefit_mode']=='consume'
assert policy['benefit_ratio_at_most']==0.97 and policy['benefit_ci_high_below']==1.0
assert policy['protected_ratio_above']==1.05 and policy['protected_ci_low_above']==1.0
assert policy['production_adoption'] is False
cases=c.read(p/'fixtures.json')
protected=[r['id'] for r in cases if int(r['id'].rsplit('-',1)[1])<=2 or r['id'].startswith('duplicate-long-')]
assert protected==policy['protected_consume_cases'] and len(protected)==18
a=c.read(p/'analysis.json')['native']['analysis']
benefits={}
for case in policy['benefit_cases']:
    row=a['paired_by_case_mode'][case+'/consume']
    benefits[case]=row['ratio_median']<=0.97 and row['bootstrap']['ci_high']<1.0
flags=[f for f in a['diagnostic_regression_flags'] if f['mode']=='consume']
vetoes=[f for f in flags if f['case'] in protected]
result={
    'schema':'litchi.performance.0805.preflight-decision.v1',
    'advance_to_workflow_trials':all(benefits.values()) and not vetoes,
    'dominant_class_benefits':benefits,
    'protected_consume_regressions':vetoes,
    'all_consume_regressions':flags,
    'production_adoption':False,
    'analysis':c.artifact(p/'analysis.json'),
    'independent_audit':c.artifact(p/'root-native-audit.json'),
}
if '--check' in sys.argv:
    assert c.read(p/'decision.json')==result
else:
    assert not (p/'decision.json').exists()
    c.write(p/'decision.json',result)
print('Preflight advance:',result['advance_to_workflow_trials'],'benefits:',benefits,'protected vetoes:',len(vetoes),'all consume flags:',len(flags))
