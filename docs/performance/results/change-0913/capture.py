"""Execute the frozen matched capture plan with create-only command directories."""
import importlib.util
import json
import os
from pathlib import Path

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('run_0913', P/'run.py')
r = importlib.util.module_from_spec(spec)
spec.loader.exec_module(r)

def main():
    r.assert_roots()
    tools_path=P/'process-tools.json'
    for tool in json.loads(tools_path.read_text())['tools']:
        path=Path(tool['executable_path'])
        assert path.stat().st_size==tool['bytes'] and r.sha256(path)==tool['sha256']
    manifest=P/'source-v1.json'
    plan_path=P/'capture-plan.json'
    r.check_manifest(manifest)
    plan=json.loads(plan_path.read_text())
    build=json.loads((P/'build-v1/command.json').read_text())
    binary=json.loads((P/'build-v1/binary.json').read_text())
    candidate=P/'binaries/candidate'
    assert build['exit_code']==0 and binary['retained']==r.descriptor(candidate)
    assert build['baseline_binary']==r.descriptor(r.BASELINE_BINARY)
    assert build['helper']==r.descriptor(P/'run.py')
    assert build['plan']==r.descriptor(plan_path)
    env={'PATH':'/usr/bin:/bin','LC_ALL':'C','LANG':'C','TZ':'UTC'}
    common={'cwd':str(r.SOURCE),'environment_overrides':env,'environment_is_complete':True,'removed_environment_keys':[], 'source':r.descriptor(manifest),'helper':r.descriptor(Path(__file__)),'build_helper':r.descriptor(P/'run.py'),'process_tools':r.descriptor(tools_path),'host_environment':r.descriptor(P/'environment-after-recovery.json'),'plan':r.descriptor(plan_path),'candidate_binary':binary['retained'],'baseline_binary':r.descriptor(r.BASELINE_BINARY)}
    for leg in plan['legs']:
        folder=P/'capture'/leg['label']
        executable=r.BASELINE_BINARY if leg['variant']=='baseline' else candidate
        argv=['/usr/bin/time','-f','max_rss_kib=%M\nuser_seconds=%U\nsystem_seconds=%S\nelapsed_seconds=%e\nmajor_faults=%F\nminor_faults=%R\nvoluntary_context_switches=%w\ninvoluntary_context_switches=%c','-o',str(folder/'process.txt'),'taskset','-c',str(plan['cpu']),str(executable),leg['route'],str(plan['warmup']),str(plan['samples']),str(folder/'artifacts')]
        r.execute(folder,argv,env,common,{'leg':leg,'variant':leg['variant'],'binary':r.descriptor(executable),'plan':r.descriptor(plan_path)})

if __name__=='__main__': main()
