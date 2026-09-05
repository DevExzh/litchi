import datetime,json,os,subprocess,time
from pathlib import Path
root=Path('/home/zhuhe/code/litchi')
out=Path('/tmp/litchi-goal-0411-checks')
env=os.environ.copy()
env.update(RUSTUP_TOOLCHAIN='1.98.1',CARGO_TARGET_DIR='/tmp/litchi-goal-0411-tests',CARGO_BUILD_JOBS='4',CARGO_INCREMENTAL='0',CARGO_PROFILE_DEV_DEBUG='0',CARGO_PROFILE_TEST_DEBUG='0')
commands=json.loads((out/'commands.json').read_text())
checks=[('harness-xls-allocator-tests',['cargo','+1.98.1','test','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--lib','--','--test-threads=4','xls_source_','allocation_metrics::tests'])]
for label,argv in checks:
 start=time.monotonic(); stamp=datetime.datetime.now(datetime.timezone.utc).isoformat()
 with (out/(label+'.log')).open('wb') as log:
  result=subprocess.run(argv,cwd=root,env=env,stdout=log,stderr=subprocess.STDOUT)
 commands.append(dict(label=label,argv=argv,cwd=str(root),started_utc=stamp,wall_seconds=time.monotonic()-start,exit_code=result.returncode))
 (out/'commands.json').write_text(json.dumps(commands,indent=2)+'\n')
 print(label,result.returncode,flush=True)
 result.check_returncode()
