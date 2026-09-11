"""Serial source-bound checks for the CFB FAT reservation change."""
import datetime,json,os,subprocess,time
from run import HERE,REPO,SCRATCH,sha,sources,write
env=dict(os.environ,CARGO_TARGET_DIR=str(SCRATCH/'target'),CARGO_BUILD_JOBS='2',CARGO_INCREMENTAL='0',CARGO_PROFILE_RELEASE_DEBUG='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',TMPDIR=str(SCRATCH),RUSTDOCFLAGS='-D warnings')
commands=[('fmt',['cargo','fmt','--all','--','--check']),('ole2-tests',['cargo','test','--release','--locked','-p','litchi-cfb','-p','litchi-xls','-p','litchi-doc','-p','litchi-ppt','--all-features','--all-targets']),('ole2-doctests',['cargo','test','--release','--locked','-p','litchi-cfb','-p','litchi-xls','-p','litchi-doc','-p','litchi-ppt','--all-features','--doc']),('cfb-no-default-tests',['cargo','test','--release','--locked','-p','litchi-cfb','--no-default-features']),('cfb-clippy',['cargo','clippy','--release','--locked','-p','litchi-cfb','--all-features','--all-targets','--','-D','warnings']),('cfb-rustdoc',['cargo','doc','--release','--locked','-p','litchi-cfb','--all-features','--no-deps'])]
for name,cmd in commands:
    current=sources();assert current==json.loads((HERE/'source-manifest.json').read_text())
    tick=time.monotonic();started=datetime.datetime.now(datetime.timezone.utc).isoformat()
    with (HERE/f'{name}.log').open('x') as out:r=subprocess.run(cmd,cwd=REPO,env=env,stdout=out,stderr=subprocess.STDOUT)
    unchanged=current==sources()
    write(HERE/f'{name}-receipt.json',{'command':cmd,'started_utc':started,'elapsed_seconds':time.monotonic()-tick,'exit_code':r.returncode,'source_unchanged':unchanged,'source_manifest_sha256':sha(HERE/'source-manifest.json'),'log_sha256':sha(HERE/f'{name}.log'),'environment':{k:env[k] for k in ['CARGO_TARGET_DIR','CARGO_BUILD_JOBS','CARGO_INCREMENTAL','CARGO_PROFILE_RELEASE_DEBUG','CARGO_PROFILE_DEV_DEBUG','CARGO_PROFILE_TEST_DEBUG','TMPDIR','RUSTDOCFLAGS']}})
    print(name,r.returncode,flush=True);assert r.returncode==0 and unchanged
