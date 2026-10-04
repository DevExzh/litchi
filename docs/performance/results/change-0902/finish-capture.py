"""Complete the frozen capture serially after a successful native stage."""
import pathlib,json,time,subprocess,sys
p=pathlib.Path(__file__).resolve().parent
f=p/'native-capture-run.json'
while not f.exists():time.sleep(1)
r=json.loads(f.read_text());assert r['exit_code']==0 and r['inputs_unchanged']
for label,helper,action in [('allocator-capture','run_v2.py','allocator'),('final-analysis','analyze_v3.py','all')]:
 subprocess.run([sys.executable,'-B',str(p/'command.py'),label,'--','python3','-B',str(p/helper),action],check=True)
