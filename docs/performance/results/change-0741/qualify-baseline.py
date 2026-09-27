"""Pre-edit baseline qualification against unchanged full 0739/0740 oracle."""
import importlib.util,json,subprocess,time
from build import P,ROOT,sha
spec=importlib.util.spec_from_file_location('contract',P/'baseline-contract.py');m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
if __name__=='__main__':
    m.guard();folder=P/'baseline-qualification';folder.mkdir();rows=[]
    for b in m.read('build.json'):
        for case in ['pptx_cross_copy_plain_lifecycle','pptx_cross_copy_media_rich_lifecycle']:
            i=len(rows);report=folder/f'{i:02d}.json';command=['taskset','-c','12',b['binary'],'--warmup','0','--samples','1','--case',case,'--json',str(report)]
            start=time.time()
            with (folder/f'{i:02d}.stdout').open('w') as out,(folder/f'{i:02d}.stderr').open('w') as err:r=subprocess.run(command,cwd=ROOT,stdout=out,stderr=err)
            row={'lane':b['lane'],'case':case,'samples':1,'warmups':0,'command':command,'started':start,'ended':time.time(),'exit':r.returncode};rows.append(row)
            (folder/'manifest.json').write_text(json.dumps(rows,indent=2)+'\n');assert r.returncode==0,row
            assert m.projection(json.loads(report.read_text()),row)==m.read('oracle.json')[case]
            row['files']={f.name:sha(f) for f in folder.glob(f'{i:02d}.*')}
            (folder/'manifest.json').write_text(json.dumps(rows,indent=2)+'\n');print(f'PASS baseline qualification {i+1}/4',flush=True)
    m.guard()
