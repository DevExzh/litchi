"""Keep validation success separate from the optimization's adoption decision."""
import sys
import driver as d
import analyze as a

def decision():
 result=a.analysis('native');plan=d.read(d.P/'plan.json');policy=plan['adoption']
 benefits=[]
 for row in result['summaries']:
  latency=row['paired']['p50_ns']
  if row['case'].startswith('fresh-') and latency['median']<=policy['minimum_material_improvement_ratio'] and latency['bootstrap_95'][1]<1:
   benefits.append(row['case'])
 flags=result['flags']
 # Every triggered regression, spread or size-growth rule requires recorded
 # review. No favorable aggregate can hide one. Review does not rewrite data.
 status='eligible' if benefits and not flags else 'review-required' if benefits else 'reject'
 return dict(validation='pass',adoption=status,material_fresh_benefits=benefits,review_flags=flags,policy=policy,scope='Prepared OPC graph to_bytes only; no filesystem-save or format lifecycle claim.')
if __name__=='__main__':
 value=decision()
 if sys.argv[1:]==['--check']:assert d.read(d.P/'decision.json')==value
 else:assert not sys.argv[1:];d.write(d.P/'decision.json',value)
 print(value)
