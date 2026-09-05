#!/usr/bin/env python3
"""Replay attribution corruptions with mutation-specific rejection checks."""
import argparse,json,shutil,subprocess,tempfile
from pathlib import Path

NAMES=['owned-profile-script.stdout','owned-profile-self.stdout','owned-profile.json','owned-profile.catalog.json','commands.json','postprocess-commands.json','protocol.json','candidate-build-identity.json']
def main():
 p=argparse.ArgumentParser();p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();bundle=Path(__file__).resolve().parents[1];records=[]
 with tempfile.TemporaryDirectory(prefix='litchi-goal-0412-attribute-negative-') as td:
  root=Path(td)
  for name in NAMES:
   src=bundle/'capture'/name;dst=root/name
   if src.exists():shutil.copy2(src,dst)
   else:dst.write_bytes(subprocess.check_output(['zstd','-q','-dc',str(src)+'.zst']))
  base=['python3',str(bundle/'attribution/attribute.py'),'--capture',str(root),'--script',str(root/NAMES[0]),'--report',str(root/NAMES[1]),'--profile',str(root/NAMES[2]),'--catalog',str(root/NAMES[3]),'--commands',str(root/NAMES[4]),'--postprocess-commands',str(root/NAMES[5]),'--protocol',str(root/NAMES[6]),'--build-identity',str(root/NAMES[7]),'--repo',str(a.repo_root.resolve())]
  raw=(root/NAMES[0]).read_text()
  cases=[('missing-protocol','protocol'),('wrong-record-graph','capture_command'),('mixed-cycles-k','sample_parser'),('zero-period','sample_parser'),('malformed-header','sample_parser'),('unparsed-frame','sample_parser'),('nonnull-source','profile_result')]
  for name,expected in cases:
   extra=[];mutant=root/(name+'.input')
   if name=='missing-protocol':extra=['--protocol',str(mutant)]
   elif name=='wrong-record-graph':
    v=json.loads((root/'commands.json').read_text())
    for r in v:
     if r['label']=='owned-profile':r['argv']=['none' if x=='fp,127' else x for x in r['argv']]
    mutant.write_text(json.dumps(v));extra=['--commands',str(mutant)]
   elif name=='nonnull-source':
    v=json.loads((root/'owned-profile.json').read_text());v['results'][0]['source']={'synthetic_counter':1};mutant.write_text(json.dumps(v));extra=['--profile',str(mutant)]
   else:
    prefix={'mixed-cycles-k':'bad 1 1.0: 5 cycles:k:\n\t0001 foo+0x0 (x)\n\n','zero-period':'bad 1 1.0: 0 cycles:u:\n\t0001 foo+0x0 (x)\n\n','malformed-header':'bad 1 1.0: NOT_A_PERIOD cycles:u:\n\n','unparsed-frame':'bad 1 1.0: 5 cycles:u:\n\tINVALID FRAME TEXT\n\n'}[name]
    mutant.write_text(prefix+raw);extra=['--script',str(mutant)]
   out=root/(name+'.output.json');argv=[*base,*extra,'--output',str(out)];r=subprocess.run(argv,text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
   v=json.loads(out.read_text()) if out.exists() else {};identity=v.get('capture_identity',{}).get('required_identity',{});failures=identity.get('failures',[])
   if r.returncode!=2 or identity.get('status')!='failed' or not any(x.startswith(expected+':') for x in failures):raise RuntimeError(f'{name}: unexpected rejection {r.returncode} {r.stdout}')
   records.append(dict(mutation=name,expected_failed_check=expected,argv=argv,exit_code=r.returncode,required_identity=identity,diagnostic=r.stdout))
 a.output.write_text(json.dumps(dict(status='passed',probes=records),indent=2)+'\n');print(f'Passed {len(records)} attribution corruption probes')
if __name__=='__main__':main()
