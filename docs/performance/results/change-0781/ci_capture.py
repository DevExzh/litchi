"""Root-only serial captures for the existing CI matrix contract."""
import subprocess,time
import custody as c
out=c.P/'ci-capture';assert not out.exists();out.mkdir()
binary=c.read(c.P/'ci-baseline-build/binary.json');assert c.artifact(binary['path'])==binary
frozen=c.source();assert frozen==c.read(c.P/'ci-baseline-build/source.json');c.write(out/'source.json',frozen)
rows=[]
for name,args in [('smoke',['--samples','2','--shape','tiny','--payload','compressible','--writer-shape','tiny','--xlsx-shape','tiny','--semantic-shape','tiny']),('full',['--samples','15','--corpus-manifest',str(out/'full.corpus-manifest-v2.json')])]:
 report=out/f'{name}.json';log=out/f'{name}.log';cmd=[binary['path'],*args,'--json',str(report)];start=time.time()
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
 row={'mode':name,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'log':c.artifact(log)}
 if report.exists():row['report']=c.artifact(report)
 catalog=out/'full.corpus-manifest-v2.json'
 if name=='full' and catalog.exists():row['catalog']=c.artifact(catalog)
 rows.append(row);c.write(out/'receipts.json',rows)
 assert r.returncode==0,row
 assert c.source()==frozen
 print(name,'CI capture complete;',len(c.read(report)['results']),'rows',flush=True)
c.write(out/'complete.json',{'processes':len(rows),'receipts':c.artifact(out/'receipts.json')})
