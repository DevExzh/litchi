"""Extract disjoint commit children and explicitly overlapping diagnostics."""
import json,re
from pathlib import Path
HERE=Path(__file__).resolve().parent
COMMIT='litchi_xlsx::workbook::edit::semantic::transaction::Edit::commit'
CHILDREN={
 'source_store':'litchi_xlsx::workbook::model::Worksheet::store',
 'rewrite':'litchi_xlsx::raw::worksheet::edit::package::rewrite',
 'validation_parse':'litchi_xlsx::raw::worksheet::parse',
 'compaction':'litchi_xlsx::raw::compact::changed_worksheet',
}
DIAGNOSTICS={
 'all_worksheet_parse':'litchi_xlsx::raw::worksheet::parse',
 'eager_parser':'litchi_xlsx::raw::worksheet::codec::<impl litchi_xlsx::raw::worksheet::model::Parser>::parse',
 'snapshot_scan':'litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::scan_with_limit',
 'snapshot_cell_address':'litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::cell_address',
 'snapshot_cell_tag':'litchi_xlsx::raw::worksheet::edit::codec::wire::cell_tag',
 'mce_processing':'litchi_ooxml_common::mce::codec::process_ooxml',
 'shared_formula_resolution':'litchi_xlsx::raw::worksheet::semantic::resolve_shared_formulas',
}
def count(line):return int(re.match(r'\s*([\d,]+)',line)[1].replace(',',''))
def analyze():
 out={}
 for lane in ['profile-r1','profile-r2']:
  total=int(re.search(r'^summary: (\d+)',(HERE/(lane+'.out')).read_text(),re.M)[1]);assert total>0
  lines=(HERE/(lane+'-inclusive.txt')).read_text().splitlines()
  heads=[i for i,l in enumerate(lines) if ' * ' in l and l.split('???:')[-1].split(' [')[0]==COMMIT]
  assert len(heads)==1
  head=heads[0];assert count(lines[head])==total
  children={};i=head+1
  while i<len(lines) and lines[i].strip():
   line=lines[i]
   for key,symbol in CHILDREN.items():
    if '>   ???:'+symbol+' (' in line:
     assert key not in children and '(6x)' in line
     n=count(line);children[key]={'ir':n,'percent':100*n/total,'direct_calls':6}
   i+=1
  assert set(children)==set(CHILDREN)
  assert sum(v['ir'] for v in children.values())<=total
  diagnostics={}
  for key,symbol in DIAGNOSTICS.items():
   matches=[l for l in lines if ' * ' in l and l.split('???:')[-1].split(' [')[0]==symbol]
   assert len(matches)==1,(lane,key)
   n=count(matches[0]);diagnostics[key]={'ir':n,'percent':100*n/total}
  out[lane]={'total_ir':total,'direct_commit_children':children,'overlapping_diagnostics':diagnostics}
 b,a=[out[k]['total_ir'] for k in ['profile-r1','profile-r2']]
 return {'profiles':out,'repeat_ir_change_percent':100*(a/b-1),'scope':'Three commit bodies only per capture, zeroed at runner entry; direct children are disjoint while named diagnostics overlap. These are simulated instruction costs, not phase clocks or cycles.','caveats':['Counter reset retains synthetic costs on active ancestors; never sum top-level exclusive rows.','Collection-off descendant calls may remain in call-count metadata. Only direct runner-to-commit (3) and named direct commit-child (6) edges are used for invocation proof.','Both Valgrind logs report brk-segment overflow and then complete successfully; instrumentation timings and RSS are excluded. No allocator/heap conclusion is drawn.'],'decision':'Prioritize the two parser traversals and snapshot layout, with operation-local allocation/phase instrumentation before fusion; shared-formula and plain-tag shortcuts alone cannot explain the major cost.'}
if __name__=='__main__':print(json.dumps(analyze(),indent=2,sort_keys=True))
