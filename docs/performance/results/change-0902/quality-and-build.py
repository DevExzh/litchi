"""Serial root quality/build queue; run only after Office tests exit."""
import pathlib,subprocess,sys,json,hashlib,shutil
p=pathlib.Path(__file__).resolve().parent
s=pathlib.Path('/tmp/litchi-0902-src');t=pathlib.Path('/home/zhuhe/code/litchi/target/litchi-0902-build')
def run(label,args):subprocess.run([sys.executable,'-B',str(p/'command.py'),label,'--',*args],check=True)
run('candidate-final-xml-tests',['cargo','test','-p','xml-minifier','--all-targets','--locked','--offline','-j','2'])
run('candidate-fmt',['cargo','fmt','-p','xml-minifier','--','--check'])
run('candidate-clippy',['cargo','clippy','-p','xml-minifier','-p','litchi-opc','-p','litchi-xlsx','-p','litchi-docx','-p','litchi-pptx','--lib','--tests','--locked','--offline','-j','2','--','-D','warnings'])
run('candidate-doc-tests',['cargo','test','-p','xml-minifier','-p','litchi-opc','-p','litchi-xlsx','-p','litchi-docx','-p','litchi-pptx','--doc','--locked','--offline','-j','2'])
run('candidate-rustdoc',['env','RUSTDOCFLAGS=-D warnings','cargo','doc','-p','xml-minifier','-p','litchi-opc','-p','litchi-xlsx','-p','litchi-docx','-p','litchi-pptx','--no-deps','--locked','--offline','-j','2'])
run('candidate-boundaries',['python3','-B','/home/zhuhe/code/litchi/tools/check_crate_boundaries.py'])
for role,binary,extra in [('native','litchi-perf-baseline',[]),('allocator','litchi-perf-baseline-alloc',['--features','allocator-metrics'])]:
 run('candidate-'+role+'-build',['cargo','build','--release','--locked','--offline','--manifest-path','tools/perf-baseline/Cargo.toml','--bin',binary,'-j','2',*extra])
 dest=p/'binaries'/('candidate-'+role);dest.parent.mkdir(exist_ok=True);assert not dest.exists();shutil.copy2(t/'release'/binary,dest);shutil.copy2(t/'release'/(binary+'.d'),p/('candidate-'+role+'.d'))
 d={'path':str(dest),'bytes':dest.stat().st_size,'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'build_receipt':str(p/('candidate-'+role+'-build-run.json'))}
 with (p/('candidate-'+role+'-binary.json')).open('x') as out:json.dump(d,out,indent=2);out.write('\n')
 print('copied',role,d['sha256'],flush=True)
