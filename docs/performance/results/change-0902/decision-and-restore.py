"""Apply the frozen rejection rule and restore only this candidate's two files."""
import pathlib,json,hashlib,subprocess,importlib.util,datetime
p=pathlib.Path(__file__).resolve().parent;root=p.parents[3]
spec=importlib.util.spec_from_file_location('analysis_0902',p/'analyze_v3.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
plan=m.load_plan();source=m.source_check(plan);m.write_once(p/'source-final.json',source)
a=m.read_json(p/'analysis.json');assert a['status']=='pass' and not a['reduction']['primary_gate_pass'];assert m.read_json(p/'raw-replay.json')['status']=='pass'
b=m.read_json(pathlib.Path(plan['source']['baseline_manifest']['path']))['files'];c=m.read_json(pathlib.Path(plan['source']['candidate_manifest']['path']))['files'];records=[]
for name in plan['source']['candidate_files']:
 f=root/name;before=m.digest(f);assert before['sha256']==c[name]['sha256']
 original=subprocess.check_output(['git','show','HEAD:'+name],cwd=root);assert len(original)==b[name]['bytes'] and hashlib.sha256(original).hexdigest()==b[name]['sha256']
 records.append((name,original,before))
for name,original,before in records:(root/name).write_bytes(original)
restored=[{'path':name,'candidate':before,'restored':m.digest(root/name)} for name,_,before in records]
assert all(r['restored']['sha256']==b[r['path']]['sha256'] for r in restored)
m.write_once(p/'decision.json',{'status':'rejected','production_retained':False,'reason':'Neither primary meets paired p50 median ratio <=0.98 and 95% upper <1.0','reduction':a['reduction'],'analysis':m.digest(p/'analysis.json'),'source_final':m.digest(p/'source-final.json'),'restored_files':restored,'recorded_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'performance_claim':'no adopted improvement; control observation does not override failed primary gate'})
print('Rejected candidate; restored exact two files to HEAD/baseline hashes')
