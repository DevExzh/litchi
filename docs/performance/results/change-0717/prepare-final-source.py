#!/usr/bin/env python3
"""Archive the unused initial build before applying final review clarifications."""
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('prepare0717',P/'custody.py')
C=importlib.util.module_from_spec(spec);spec.loader.exec_module(C)
def write(path,value):path.write_text(json.dumps(value,indent=2)+'\n')
def main():
    builds=json.loads((P/'builds.json').read_text())
    assert set(builds)=={'native','procfs'} and all(r['exit_code']==0 for r in builds.values())
    archive=P/'build-attempts/01';archive.mkdir(parents=True)
    for name in ['builds.json','build-native.log','build-procfs.log','source.json','source.patch','source-preparation.json','build.py']:
        shutil.copy2(P/name,archive/name)
    witnesses=[]
    for row in builds.values():
        binary=row['binary'];path=Path(binary['path'])
        assert C.sha(path)==binary['sha256'] and path.stat().st_size==binary['bytes']
        path.unlink();witnesses.append(binary)
    C.BIN.rmdir()
    write(archive/'archive-removal.json',dict(reason='Unused initial build predates final scope wording and Clippy cleanup; no captures used these binaries.',binaries=witnesses,binaries_removed=True))
    for name in ['builds.json','build-native.log','build-procfs.log']:(P/name).unlink()
    path=C.ROOT/'tools/perf-baseline/src/ordinary_save.rs';s=path.read_text()
    s=s.replace('ordinary_save_phase_interval_including_before_after_procfs_snapshot_probe_overhead','ordinary_save_phase_interval_same_process_counters_including_procfs_probe_overhead')
    s=s.replace('return instrumentation_identity_for(allocation_metrics::instrumentation_identity());','instrumentation_identity_for(allocation_metrics::instrumentation_identity())')
    s=s.replace('/// warmup. They document the probe\'s own cost; they are retained as evidence\n/// and are never subtracted from an operation delta.',"/// warmup. They retain counter-delta overhead only; no control durations are\n/// measured and no control is subtracted from an operation delta. Counters\n/// describe the same process, not exclusively the operation owner.")
    s=s.replace('probe.scope.contains("procfs_snapshot_probe_overhead")','probe.scope.contains("procfs_probe_overhead")')
    path.write_text(s)
    subprocess.run(['cargo','fmt','--manifest-path','tools/perf-baseline/Cargo.toml'],cwd=C.ROOT,check=True)
    write(P/'source.json',C.census())
    prep=json.loads((P/'source-preparation.json').read_text())
    (P/'source.patch').write_bytes(subprocess.check_output(['git','diff','--',*prep['changed_files']],cwd=C.ROOT))
    prep.update(source_sha256=C.sha(P/'source.json'),patch_sha256=C.sha(P/'source.patch'))
    write(P/'source-preparation.json',prep)
    print('Archived initial builds, removed unused binaries, applied review clarifications.')
if __name__=='__main__':main()
