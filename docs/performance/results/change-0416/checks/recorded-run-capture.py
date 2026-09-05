import pathlib,subprocess,sys,json
repo=pathlib.Path('/home/zhuhe/code/litchi');root=repo/'docs/performance/results/change-0416'
samples=int(sys.argv[1]) if len(sys.argv)>1 else 300
output=root if samples==300 else root/'followup'
args=['python3',str(root/'capture.py'),'--control','/tmp/litchi-goal-0416-control-probe','--candidate','/tmp/litchi-goal-0416-candidate-probe','--control-tree','/tmp/litchi-goal-0416-control','--candidate-tree','/tmp/litchi-goal-0416-candidate','--output',str(output),'--samples',str(samples),'--warmups',str(samples//10),'--cpu','2']
for label,name in [('zip32-tiny','zip32-signed-store-deflate.zip'),('zip32-many256','many-small-zip32-signed.zip'),('central64-tiny','zip64-central-local-zip32-tail-signed.zip'),('central64-many256','many-small-central-local-zip32-tail-signed.zip')]:args+=['--fixture',label+'='+str(root/'corpus'/name)]
for p in sorted((root/'corpus').glob('*local-only*.zip')):args+=['--capability-fixture',p.stem+'='+str(p)]
for item in json.loads((root/'native-fixtures.json').read_text())['files']:args+=['--indexed-fixture','native-'+pathlib.Path(item['path']).suffix[1:]+'='+str(repo/item['path'])]
r=subprocess.run(args);r.check_returncode()
subprocess.run(['python3',str(root/'summarize.py'),'--root',str(output),'--samples',str(samples),'--warmups',str(samples//10),'--output',str(output/'guard-summary.json')],check=True)
