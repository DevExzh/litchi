from pathlib import Path
import argparse,hashlib,json,statistics,re
parser=argparse.ArgumentParser(description='Reduce all retained PPT 0362f attempts without combining different production candidates.')
parser.add_argument('--coord',type=Path,default=Path(__file__).resolve().parent)
parser.add_argument('--out',type=Path)
args=parser.parse_args()
coord=args.coord.resolve()
out=args.out or coord/'ppt-review-0362f'; out.mkdir(parents=True,exist_ok=True)
def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def artifact(p): return {'path':str(p),'bytes':p.stat().st_size,'sha256':digest(p)}
runs=[]
for pattern in ['ppt-phase-original-8f-v*','ppt-phase-sip-8f-v*','ppt-phase-original-final-v*','ppt-phase-sip-final-v*']:
 for directory in sorted(coord.glob(pattern)):
  if not directory.is_dir(): continue
  source=directory/'measurements.json'
  if not source.exists(): continue
  raw=json.loads(source.read_text())
  files=[artifact(p) for p in sorted(directory.rglob('*')) if p.is_file()]
  inventory=out/(directory.name+'-artifacts.json')
  inventory.write_text(json.dumps(files,indent=2)+'\n')
  abort=directory/'abort.json'
  runs.append({'name':directory.name,'raw_measurements':artifact(source),'artifact_inventory':artifact(inventory),
   'raw_total_bytes':sum(p['bytes'] for p in files),'disposition':json.loads(abort.read_text()) if abort.exists() else raw['status'],
   'before_commit':raw['before_commit'],'after_commit':raw['after_commit'],'binaries':raw['binaries'],
   'runner':raw['runner'],'harness':raw['harness'],'fixture':raw['fixture'],'command':raw['command'],
   'cpu':raw['cpu'],'minimum_phase_ns':raw['minimum_phase_ns'],'calibration_target_ns':raw['calibration_target_ns'],
   'calibration':raw['calibration'],'counts':raw.get('counts'),'paired':raw['paired']})
comparisons={}
for group in ['original-8f','sip-8f','original-final','sip-final']:
 selected=[r for r in runs if r['name'].startswith('ppt-phase-'+group+'-')]
 blocks=[b for r in selected for b in r['paired']]
 if not blocks: continue
 for b in blocks:
  assert len(b['legs'])==4
  assert min(s for leg in b['legs'] for s in leg['spans_ns'].values())>=150000000
 comparisons[group]={'blocks':len(blocks),'required_blocks':5,'status':'coordinator review pending' if len(blocks)>=5 else 'incomplete',
 'median_after_before':{k:statistics.median(b['after_before'][k] for b in blocks) for k in blocks[0]['after_before']},
 'block_after_before':[b['after_before'] for b in blocks],
 'minimum_phase_ns':min(s for b in blocks for leg in b['legs'] for s in leg['spans_ns'].values())}
comparisons['original-final']['status']='five qualified blocks; improvements do not waive the matched-Sip replay regression'
comparisons['sip-final']['status']='returned for repair: forward and inverse replay exceed reference in four of five paired blocks'
summary={'status':'returned for replay repair; no performance exception accepted','preserved_candidate_commit':'be72b3a163f4fc1bc58a54c4d9a18b7fb8eba0b6','selection':'All completed guarded ABBA blocks retained; incomplete or overlapped blocks excluded regardless of values. Independent calibration per attempt; per-operation timings normalize counts.','comparisons':comparisons,'runs':runs,'final_candidate_build':json.loads((coord/'ppt-final-admission-20261021/manifest.json').read_text()),'matched_sip_build':json.loads((coord/'ppt-sip-common-lock-20261021/manifest.json').read_text()),'runner_checks':json.loads((coord/'ppt-phase-runner-checks/final-validation.json').read_text())}
(out/'measurements.json').write_text(json.dumps(summary,indent=2)+'\n')
logs=[]
for p in sorted((coord/'ppt-integration-0362f').rglob('*.log')):
 s=p.read_text(); matches=re.findall(r'test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored;',s)
 logs.append({'artifact':artifact(p),'test_targets':[{'status':a,'passed':int(b),'failed':int(c),'ignored':int(d)} for a,b,c,d in matches]})
checks=[{'path':str(p),'receipts':json.loads(p.read_text())} for p in sorted((coord/'ppt-integration-0362f').rglob('checks.json'))]
(out/'validation.json').write_text(json.dumps({'logs':logs,'checks':checks,'library_exit':(coord/'ppt-integration-0362f/lib-tests.exit').read_text().strip(),'distinct_tests':2107,'ignored':3,'scope':'The broad libraries and integration targets passed before the final early-length admission fix. After that fix, all 15 scoped tests, two-crate all-target/all-feature Clippy and workspace formatting passed. Full integrated gate remains required.','note':'The added tight-profile test initially expected construction to succeed; it correctly failed admission at open. Corrected expectation and separate large-directory scan boundary pass. Initial failed receipt retained.'},indent=2)+'\n')
print(json.dumps({'comparisons':comparisons,'output':str(out),'summary_bytes':(out/'measurements.json').stat().st_size},indent=2))
