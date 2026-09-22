"""Serial source-bound rebuild of existing native/allocator lifecycle harness."""
import datetime, hashlib, json, os, pathlib, platform, subprocess, time
P = pathlib.Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0740'

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def write(name, value):
    (P/name).write_text(json.dumps(value, indent=2)+'\n')

def census():
    paths = subprocess.check_output(['git','ls-files','-z'],cwd=ROOT).decode().split('\0')
    return {n:sha(ROOT/n) for n in paths if n and (n.startswith('crates/') or n.startswith('tools/perf-baseline/')) and (n.endswith('.rs') or pathlib.Path(n).name in ['Cargo.toml','Cargo.lock'])}

if __name__ == '__main__':
    assert not (P/'build.json').exists() and not TARGET.exists()
    initial = census()
    write('source.json', {'head':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'files':initial})
    write('constraints.json',{str(p.relative_to(ROOT)):sha(p) for p in [ROOT/'docs/GOAL.md',ROOT/'docs/CRUD_Scenario_Checklist.md',*sorted((ROOT/'docs/adr').glob('*.md'))]})
    write('environment.json',{'utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'platform':platform.platform(),'cpuinfo':pathlib.Path('/proc/cpuinfo').read_text().split('\n\n')[0],'meminfo':pathlib.Path('/proc/meminfo').read_text(),'affinity':sorted(os.sched_getaffinity(0)),'rustc':subprocess.check_output(['rustc','-Vv'],text=True),'cargo':subprocess.check_output(['cargo','-V'],text=True),'environment':{k:os.environ.get(k) for k in ['RUSTFLAGS','CARGO_ENCODED_RUSTFLAGS','RUSTUP_TOOLCHAIN','LD_PRELOAD','MALLOC_ARENA_MAX','MALLOC_MMAP_THRESHOLD_','MALLOC_TRIM_THRESHOLD_','GLIBC_TUNABLES']}})
    env=os.environ|{'CARGO_TARGET_DIR':str(TARGET),'CARGO_BUILD_JOBS':'2'}
    results=[]
    for lane, extra in [('native',[])]:
        binary='litchi-perf-baseline'+('-alloc' if extra else '')
        cmd=['cargo','build','--release','--offline','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin',binary,*extra]
        start=time.time()
        with (P/f'build-{lane}.log').open('w') as log:
            result=subprocess.run(cmd,cwd=ROOT,env=env,stdout=log,stderr=subprocess.STDOUT)
        row={'lane':lane,'command':cmd,'started':start,'ended':time.time(),'exit':result.returncode,'log_sha256':sha(P/f'build-{lane}.log')}
        if result.returncode==0:
            exe=TARGET/'release'/binary
            row|={'binary':str(exe),'binary_sha256':sha(exe),'binary_bytes':exe.stat().st_size}
        results.append(row);write('build.json',results)
        assert result.returncode==0,row
        assert census()==initial
    print('PASS one serial source-bound native harness build',flush=True)
