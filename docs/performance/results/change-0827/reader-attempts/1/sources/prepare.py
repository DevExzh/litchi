"""Prepare 0827 source archives and independently refreshed host inputs."""
import hashlib,json,os,shutil,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
OLD=P.parent/'change-0824'
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,v):
    assert not p.exists(),p
    p.write_text(json.dumps(v,indent=2,sort_keys=True)+'\n')
def main():
    base=subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT).decode().strip()
    assert base.startswith('c5d375083c')
    assert not (ROOT.parent/'litchi-target-0827').exists()
    assert not (ROOT.parent/'litchi-fs-0827').exists()
    arch=json.loads((OLD/'architecture-inputs.json').read_text())
    assert all(sha(ROOT/n)==h for n,h in arch.items())
    write(P/'architecture-inputs.json',arch)
    origin=json.loads((OLD/'origin.json').read_text())
    assert all(sha(ROOT/n)==h for n,h in origin['unrelated'].items())
    origin.update(schema='litchi.performance.0827.origin.v1',base=base,target=str(ROOT.parent/'litchi-target-0827'),scratch=str(ROOT.parent/'litchi-fs-0827'),production_changed_at_freeze=False,runtime_harness_changed=False)
    write(P/'origin.json',origin)
    for leg in ('before','after'):
        dst=P/'candidate'/leg;dst.mkdir(parents=True)
        for name in ('transaction.rs','xml.rs'):
            source=OLD/'candidate'/leg/name
            if leg=='after': assert source.read_bytes()==(ROOT/'crates/litchi-pptx/src/opened'/name).read_bytes()
            shutil.copy2(source,dst/name)
    shutil.copytree(OLD/'inputs',P/'inputs')
    write(P/'root-inputs.json',json.loads((OLD/'root-inputs.json').read_text().replace('0824','0827')))
    locks=json.loads((OLD/'lock-parity.json').read_text());locks.pop('probe_locks');locks['schema']='litchi.performance.0827.lock-parity.v1';write(P/'lock-parity.json',locks)
    for name in ('corpus-inputs.json','provenance.json'):
        val=json.loads((P.parent/'change-0819'/name).read_text().replace('litchi.performance.0819','litchi.performance.0827'))
        write(P/name,val)
    for name,desc in json.loads((P/'corpus-inputs.json').read_text()).items():
        assert sha(ROOT/name)==desc['sha256'] and (ROOT/name).stat().st_size==desc['bytes']
    host=json.loads((OLD/'host.json').read_text());host.update(schema='litchi.performance.0827.host.v1',uname=subprocess.check_output(['uname','-a']).decode().strip(),target=str(ROOT.parent/'litchi-target-0827'),scratch=str(ROOT.parent/'litchi-fs-0827'),affinity_available=sorted(os.sched_getaffinity(0)),affinity_selected=[12],logical_cpu_count=os.cpu_count())
    host['mem_total']=next(x for x in Path('/proc/meminfo').read_text().splitlines() if x.startswith('MemTotal:'))
    host['mount']=subprocess.check_output(['findmnt','-n','-o','SOURCE,FSTYPE,OPTIONS','-T',str(ROOT)]).decode().strip()
    host['cgroup']={str(q):q.read_text().strip() for n in ('cpu.max','memory.max','cpuset.cpus.effective') if (q:=Path('/sys/fs/cgroup')/n).is_file()}
    subprocess.run(['taskset','-c','12','true'],check=True);write(P/'host.json',host)
    tc={'schema':'litchi.performance.0827.toolchain.v1'}
    for cmd in ('cargo -V','rustc -Vv','python3 --version','git --version','perf --version'):
        tc[cmd]=subprocess.check_output(cmd.split()).decode().strip()
    write(P/'toolchain.json',tc)
    print('0827 preparation PASS: current HEAD, source archives,35 normative inputs,3 real corpora,locks,provenance,host,unrelated files')
if __name__=='__main__': main()
