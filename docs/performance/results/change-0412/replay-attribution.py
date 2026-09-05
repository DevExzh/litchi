#!/usr/bin/env python3
"""Materialize lossless profile inputs privately and replay CPU attribution."""
import argparse,shutil,subprocess,tempfile
from pathlib import Path

def main():
 p=argparse.ArgumentParser();p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args();bundle=Path(__file__).resolve().parent
 with tempfile.TemporaryDirectory(prefix='litchi-goal-0412-attribution-replay-') as td:
  capture=Path(td)
  for name in ['owned-profile-script.stdout','owned-profile-self.stdout','owned-profile.json','owned-profile.catalog.json','commands.json','postprocess-commands.json','protocol.json','candidate-build-identity.json']:
   src=bundle/'capture'/name;dst=capture/name
   if src.exists():shutil.copy2(src,dst)
   else:dst.write_bytes(subprocess.check_output(['zstd','-q','-dc',str(src)+'.zst']))
  argv=['python3',str(bundle/'attribution/attribute.py'),'--capture',str(capture),'--script',str(capture/'owned-profile-script.stdout'),'--report',str(capture/'owned-profile-self.stdout'),'--profile',str(capture/'owned-profile.json'),'--catalog',str(capture/'owned-profile.catalog.json'),'--commands',str(capture/'commands.json'),'--postprocess-commands',str(capture/'postprocess-commands.json'),'--protocol',str(capture/'protocol.json'),'--build-identity',str(capture/'candidate-build-identity.json'),'--repo',str(a.repo_root.resolve()),'--output',str(a.output.resolve())]
  subprocess.run(argv,check=True)
if __name__=='__main__':main()
