"""Replay retained raw evidence without requiring the cleaned live source roots."""
import pathlib,hashlib,importlib.util,json,tarfile,argparse
p=pathlib.Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('reduction_0902',p/'analyze_v3.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
a=argparse.ArgumentParser();a.add_argument('--output',type=pathlib.Path,required=True);args=a.parse_args()
plan=m.load_plan();m.case_rows(plan)
bp=m.descriptor(plan['source']['baseline_manifest'],'baseline manifest');cp=m.descriptor(plan['source']['candidate_manifest'],'candidate manifest')
baseline=m.read_json(pathlib.Path(bp['path']));candidate=m.read_json(pathlib.Path(cp['path']))
b=baseline['files'];c=candidate['files'];assert set(b)==set(c) and len(b)==10771
names=set(plan['source']['candidate_files']);assert len(names)==2
changed={n for n in b if (b[n]['bytes'],b[n]['sha256'])!=(c[n]['bytes'],c[n]['sha256'])};assert changed==names
archive=m.descriptor(baseline['archive'],'baseline source archive');seen=set()
with tarfile.open(archive['path'],'r|gz') as tar:
 for member in tar:
  if member.name in b:
   raw=tar.extractfile(member).read();d=b[member.name]
   assert len(raw)==d['bytes'] and hashlib.sha256(raw).hexdigest()==d['sha256'],member.name
   assert member.name not in seen;seen.add(member.name)
assert seen==set(b)
for name in names:
 actual=m.digest(p/'candidate-overlay'/name);assert actual['bytes']==c[name]['bytes'] and actual['sha256']==c[name]['sha256']
binaries=m.binaries_check(plan);validator=m.load_validator(plan)
rows={stage:m.capture_review(plan,stage,binaries,validator) for stage in ['qualification','native','allocator']}
m.compare_identities(sum(rows.values(),[]));reduction=m.native_reduce(rows['native'],plan)
canonical=m.read_json(p/'analysis.json')
assert canonical['status']=='pass' and canonical['stage']=='all'
for stage in rows:assert rows[stage]==canonical[stage],stage
assert reduction==canonical['reduction']
m.write_once(args.output,{'status':'pass','mode':'retained artifact/semantic/numeric replay; no live source-root claim','baseline_archive':archive,'baseline_source_files_verified':len(seen),'candidate_override_files_verified':sorted(names),'children':{k:len(v) for k,v in rows.items()},'analysis':m.digest(p/'analysis.json'),'analyzer':m.digest(p/'analyze_v3.py'),'replay_helper':m.digest(pathlib.Path(__file__)),'reduction':reduction})
print('historical replay PASS')
