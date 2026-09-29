"""Select concrete inclusive symbol costs without adding overlapping regions."""
import driver as d
TARGETS={
 'topology':'litchi_opc::source_backed::SourceBackedPackage::prepare_topology_state',
 'part_name_insert':'litchi_opc::members::PartNameIndex::insert',
 'topology_limits':'litchi_opc::source_backed::SourceBackedPackage::validate_topology_limits',
 'physical_lookup':'litchi_opc::source_backed::SourceBackedPackage::build_physical_member_lookup',
 'index':'soapberry_zip::preserve::PreservationIndex<R>::new_with_limits_and_policy',
 'write_to':'soapberry_zip::preserve::PreservationIndex<R>::write_to',
 'layout':'soapberry_zip::preserve::PreservationIndex<R>::validate_output_layout',
 'prepare':'soapberry_zip::preserve::PreservationIndex<R>::prepare_with',
}
def verify():
 a=d.read(d.P/'analysis.json');rows=[]
 for x in a['profiles']:
  p=x['profile'];symbols={}
  for key,name in TARGETS.items():
   found=[f for f in p['function_rows'] if f['name']==name];assert len(found)==1,(key,found)
   f=found[0];cost=f['self_ir']+f['direct_children_ir'];assert 0<=cost<=p['summary_ir']
   symbols[key]=dict(name=name,self_ir=f['self_ir'],inclusive_ir=cost,owner_fraction=cost/p['summary_ir'])
  rows.append(dict(row=x['row'],owner_ir=p['summary_ir'],symbols=symbols))
 return dict(status='pass',rows=rows,analysis=d.desc(d.P/'analysis.json'),reducer=d.desc(d.P/'attribution.py'),scope='Inclusive costs overlap; topology contains insert/limits/lookup, and write_to contains prepare/layout. Do not sum these rows or infer elapsed phase times.')
if __name__=='__main__':
 value=verify();p=d.P/'attribution.json'
 if p.exists():assert d.read(p)==value
 else:d.write(p,value)
 print('Attribution PASS',len(value['rows']),'profiles')
