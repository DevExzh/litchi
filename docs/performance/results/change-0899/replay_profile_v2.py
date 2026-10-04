#!/usr/bin/env python3
"""Replay retained semantics, explicitly accounting for data-file mode drift."""
from pathlib import Path
import hashlib, importlib.util, json, sys
P=Path(__file__).resolve().parent
OLD=P.parent/'change-0892'
sys.path.insert(0,str(OLD))
spec=importlib.util.spec_from_file_location('profile0892',OLD/'verify.py')
m=importlib.util.module_from_spec(spec);sys.modules[spec.name]=m;spec.loader.exec_module(m)
def d(p):return {'path':str(p.resolve()),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def main():
 output=P/'profile-content-replay.json';assert not output.exists()
 allowed={str((OLD/f'callgrind/callgrind-{i:02}/callgrind.out').resolve()) for i in range(2)}
 drift=[]
 def check_artifacts(folder,command):
  for name,expected in command.get('artifacts',{}).items():
   path=folder/name;assert path.is_file() and not path.is_symlink(),path
   actual=m.descriptor(path)
   for key in ('bytes','sha256','executable'):assert expected[key]==actual[key],(path,key)
   if expected['mode']!=actual['mode']:
    assert str(path.resolve()) in allowed and not expected['executable'] and not actual['executable'],path
    assert expected['mode']==0o600 and actual['mode']==0o664,(path,expected,actual)
    drift.append({'expected':expected,'actual':actual})
 m.check_artifacts=check_artifacts
 def no_intermediate_write(path,value):assert Path(path)==output
 m.write_once=no_intermediate_write
 result=m.verify_packet(OLD/'plan.json',OLD/'capture',OLD/'callgrind',output,OLD/'self-analysis.json',OLD/'callgrind-analysis.json')
 assert len(drift)==2 and result['status']=='pass'
 value={'schema':'litchi.profile-content-custody-0899.v1','status':'pass','scope':'Byte/hash, executable-bit, command, source, corpus and semantic replay. Two non-executable raw Callgrind data files have permission-only drift; original strict mode check remains failed. No historical artifact was changed.','strict_historical_verifier_pass':False,'content_semantic_replay_pass':True,'data_permission_drift':drift,'replay':result,'helper':d(Path(__file__)),'retained_verifier':d(OLD/'verify.py'),'workload_launched':False,'profiler_launched':False}
 with output.open('x') as f:json.dump(value,f,indent=2,sort_keys=True);f.write('\n')
 print('Content and semantic replay PASS; two data permission differences retained')
if __name__=='__main__':main()
