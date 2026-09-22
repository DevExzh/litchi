"""Qualify unchanged native outputs and perf callchain observability separately."""
import json,subprocess,time
from contract import P, ROOT, read, guard, sha, projection
CASES=['pptx_cross_copy_plain_lifecycle','pptx_cross_copy_media_rich_lifecycle']

def run():
    guard();freeze=read('freeze.json');assert all(sha(P/n)==h for n,h in freeze['scripts'].items());folder=P/'formal';folder.mkdir();receipts=[]
    binary=read('build-fp.json')['binary']
    for index,(mode,case,samples) in enumerate([('profile',c,100 if 'plain' in c else 10) for repeat in range(3) for c in (CASES if repeat%2==0 else CASES[::-1])]):
        report=folder/f'{index:02d}.json';base=['taskset','-c','12',binary,'--warmup','0','--samples',str(samples),'--case',case,'--json',str(report)]
        command=base if mode=='native' else ['perf','record','-e','cycles:u','-F','997','--call-graph','fp,127','-o',str(folder/f'{index:02d}.perf.data'),'--',*base]
        start=time.time()
        with (folder/f'{index:02d}.stdout').open('w') as out,(folder/f'{index:02d}.stderr').open('w') as err:
            r=subprocess.run(command,cwd=ROOT,stdout=out,stderr=err)
        receipt={'mode':mode,'case':case,'samples':samples,'warmups':0,'command':command,'started':start,'ended':time.time(),'exit':r.returncode};receipts.append(receipt)
        (folder/'manifest.json').write_text(json.dumps(receipts,indent=2)+'\n')
        assert r.returncode==0,receipt
        if mode=='profile':
            errors=(folder/f'{index:02d}.stderr').read_text().lower()
            assert not any(word in errors for word in ['failed','error:','corruption','lost ']),errors
        row={'case':case,'lane':'native','build':'fp','samples':samples,'warmups':0}
        assert projection(json.loads(report.read_text()),row)==read('oracle.json')[case]
        if mode=='profile':
            cmd=['perf','script','--show-lost-events','-i',str(folder/f'{index:02d}.perf.data'),'-F','comm,pid,tid,time,event,period,ip,sym,dso']
            with (folder/f'{index:02d}.stacks').open('w') as out,(folder/f'{index:02d}.script.stderr').open('w') as err:
                r=subprocess.run(cmd,cwd=ROOT,stdout=out,stderr=err)
            receipt['script_command']=cmd;receipt['script_exit']=r.returncode;assert r.returncode==0
        receipt['files']={f.name:sha(f) for f in folder.glob(f'{index:02d}.*')}
        (folder/'manifest.json').write_text(json.dumps(receipts,indent=2)+'\n')
        print(f'PASS formal {index+1}/6 {mode} {case}',flush=True)
    guard()
if __name__=='__main__':run()
