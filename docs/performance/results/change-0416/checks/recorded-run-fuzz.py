import json,os,subprocess,pathlib
root=pathlib.Path('/home/zhuhe/code/litchi')
f=json.loads((root/'docs/performance/results/change-0416/fuzz.json').read_text())
env=os.environ|{'CARGO_BUILD_JOBS':'4','CARGO_TARGET_DIR':'/tmp/litchi-goal-0416-fuzz-target','RUSTC_BOOTSTRAP':'1','RUSTFLAGS':f['rustflags']}
for crate,kind,target in [('soapberry-zip','zip','parse_zip'),('litchi-opc','opc','parse_opc')]:
 for name,cmd in [(kind+'-fuzz-build',['cargo','+1.98.1','build','--release','--locked','--manifest-path',f'crates/{crate}/fuzz/Cargo.toml','--target',f['target'],'--bin',target]),(kind+'-fuzz',['/tmp/litchi-goal-0416-fuzz-target/'+f['target']+'/release/'+target,'/tmp/litchi-goal-0416-fuzz/'+kind,'-runs=1000','-seed=416','-max_len=1048576','-timeout=10'])]:
  result=subprocess.run(['python3','/tmp/litchi-goal-0416-run-check.py',name,*cmd],cwd=root,env=env)
  if result.returncode: raise SystemExit(result.returncode)
