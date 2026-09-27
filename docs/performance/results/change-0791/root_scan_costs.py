"""Additional raw function-row cross-check; no native workload or speedup claim."""
import importlib.util,json,sys
from pathlib import Path
import custody as c
p=c.P;ref=c.read(p/'inheritance.json')['references']['profile_analysis.py'];source=p/ref['path']
assert c.sha(source)==ref['sha256'] and source.stat().st_size==ref['bytes']
spec=importlib.util.spec_from_file_location('historical_callgrind_reader',source);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);m.HERE=p
rows=[]
for f in sorted((p/'profiles').glob('*.callgrind.1')):
 d=m.parse_raw(f);out={}
 for part in ['scan_processed_xml','inspect_element','IterState::next','CheckedAttributes as','notes::resolved','from_utf8','load_index_with_slide_root_proofs']:
  found=[v for v in d['functions'].values() if part in v['name']]
  out[part]=[{'name':v['name'],'self_ir':v['self_ir'],'child_ir':sum(e['inclusive_ir'] for e in v['edges']),'incoming_calls':sum(e['calls'] for e in m.incoming_edges(d,v['id'])), 'direct_children':[{**e, 'callee_name':d['functions'][e['callee_id']]['name']} for e in v['edges']] if part=='inspect_element' else []} for v in found]
 total=sum(v['self_ir'] for v in d['functions'].values());assert total==d['header']['summary_ir']
 rows.append({'profile':f.name,'summary_ir':d['header']['summary_ir'],'self_sum':total,'selected_functions':out})
assert len(rows)==6
result={'scope':'Root cross-check of raw profile function rows via hash-bound0784 parser; guest instructions only, nested child costs overlap','rows':rows}
path=p/'root-scan-costs.json'
if '--check' in sys.argv:assert c.read(path)==result;print('Raw function-row cross-check PASS')
else:assert not path.exists();c.write(path,result);print(json.dumps(rows[0],indent=2))
