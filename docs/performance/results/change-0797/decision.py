"""Frozen micro-preflight advancement only, never production adoption."""
import sys
import custody as c
p=c.P;policy=c.read(p/'plan.json')['preflight_decision']
assert policy=={'advance_requires_semantics':True,'advance_requires_distinct_1_consume_ratio_at_most':0.97,'advance_requires_distinct_1_consume_ci_high_below':1.0,'advance_requires_zero_consume_diagnostic_regressions':True,'construct_flags_are_diagnostic_only':True,'failure_action':'archive candidate and retain production baseline; do not run public-workflow adoption trials for this candidate','success_action':'candidate eligible only for fresh public-workflow/resource/cross-format trials; no production adoption in this packet'}
a=c.read(p/'analysis.json')['native']['analysis'];r=a['paired_by_case_mode']['distinct-1/consume']
flags=[f for f in a['diagnostic_regression_flags'] if f['mode']=='consume'];benefit=r['ratio_median']<=0.97 and r['bootstrap']['ci_high']<1
result={'schema':'litchi.performance.0797.preflight-decision.v1','advance_to_workflow_trials':benefit and not flags,'single_attribute_benefit':benefit,'consume_regressions':flags,'production_adoption':False,'analysis':c.artifact(p/'analysis.json'),'independent_audit':c.artifact(p/'root-native-audit.json')}
if '--check' in sys.argv:assert c.read(p/'decision.json')==result
else:assert not (p/'decision.json').exists();c.write(p/'decision.json',result)
print('Preflight advance:',result['advance_to_workflow_trials'],'single-attribute benefit:',benefit,'consume regression rows:',len(flags))
