"""Build/run the pre-change durable-patch fixture generator serially."""
import json,os,subprocess,time,tomllib
from build import P,ROOT,TARGET,census,sha
if __name__=='__main__':
    assert census()==json.loads((P/'source.json').read_text())['files']
    manifest=P/'legacy-fixture'/'Cargo.toml';output=P/'legacy-fixture'/'captured';assert not output.exists()
    env=os.environ|{'CARGO_TARGET_DIR':str(TARGET),'CARGO_BUILD_JOBS':'2'}
    commands=[['cargo','generate-lockfile','--offline','--manifest-path',str(manifest)],['cargo','run','--release','--offline','--locked','--manifest-path',str(manifest),'--',str(output)]]
    runs=[]
    for i,command in enumerate(commands):
        log=P/f'legacy-capture-{i}.log';start=time.time()
        with log.open('w') as out:r=subprocess.run(command,cwd=ROOT,env=env,stdout=out,stderr=subprocess.STDOUT)
        runs.append({'command':command,'exit':r.returncode,'started':start,'ended':time.time(),'log':log.name,'sha256':sha(log)})
        (P/'legacy-capture-attempt.json').write_text(json.dumps(runs,indent=2)+'\n');assert r.returncode==0
    binary=TARGET/'release'/tomllib.loads(manifest.read_text())['package']['name']
    assert census()==json.loads((P/'source.json').read_text())['files']
    for b in json.loads((P/'build.json').read_text()):assert sha(__import__('pathlib').Path(b['binary']))==b['binary_sha256']
    result={'runs':runs,'binary':str(binary),'binary_sha256':sha(binary),'binary_bytes':binary.stat().st_size,'source_sha256':sha(P/'source.json'),'generator':{str(f.relative_to(P)):sha(f) for f in manifest.parent.rglob('*') if f.is_file() and not f.is_relative_to(output)},'captured':{str(f.relative_to(P)):sha(f) for f in output.rglob('*') if f.is_file()}}
    assert result['captured'];(P/'legacy-capture.json').write_text(json.dumps(result,indent=2)+'\n');print('PASS pre-change durable fixture captured; baseline binaries/source unchanged')
