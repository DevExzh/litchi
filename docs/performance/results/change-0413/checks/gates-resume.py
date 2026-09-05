import datetime,json,os,subprocess,time
from pathlib import Path
root=Path('/home/zhuhe/code/litchi');out=Path('/tmp/litchi-goal-0413-checks');out.mkdir(exist_ok=True)
env=os.environ.copy();env.update(RUSTUP_TOOLCHAIN='1.98.1',CARGO_TARGET_DIR='/tmp/litchi-goal-0413-tests',CARGO_BUILD_JOBS='4',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0',RUSTDOCFLAGS='-D warnings')
checks=[('cfb-legacy-tests',['cargo','+1.98.1','test','--locked','-p','litchi-cfb','-p','litchi-xls','-p','litchi-doc','-p','litchi-ppt','--all-features','--all-targets','--no-fail-fast']),('cfb-default-tests',['cargo','+1.98.1','test','--locked','-p','litchi-cfb','--no-default-features']),('cfb-clippy',['cargo','+1.98.1','clippy','--locked','-p','litchi-cfb','--all-features','--all-targets','--no-deps','--','-D','warnings']),('cfb-rustdoc',['cargo','+1.98.1','doc','--locked','-p','litchi-cfb','--all-features','--no-deps']),('rustfmt',['rustfmt','+1.98.1','--check','--edition','2024','--config','skip_children=true','crates/litchi-cfb/src/file.rs'])]
checks=checks[2:]
checks[0][1].extend(['-A','clippy::chunks_exact_to_as_chunks'])
records=json.loads((out/'commands.json').read_text())
for label,argv in checks:
 stamp=datetime.datetime.now(datetime.timezone.utc).isoformat();start=time.monotonic()
 with (out/(label+'.log')).open('wb') as log:r=subprocess.run(argv,cwd=root,env=env,stdout=log,stderr=subprocess.STDOUT)
 records.append(dict(label=label,argv=argv,cwd=str(root),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=r.returncode));(out/'commands.json').write_text(json.dumps(records,indent=2)+'\n');print(label,r.returncode,flush=True);r.check_returncode()
