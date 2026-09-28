"""Prepare fresh immutable inputs; never copy measurements from an earlier trial."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / "change-0823"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, value):
    assert not path.exists(), path
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def main():
    base = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT).decode().strip()
    assert base.startswith("7b268927bf")
    assert not (ROOT.parent / "litchi-target-0824").exists()
    architecture = json.loads((OLD / "architecture-inputs.json").read_text())
    assert all(sha(ROOT / name) == digest for name, digest in architecture.items())
    write(P / "architecture-inputs.json", architecture)
    origin = json.loads((OLD / "origin.json").read_text())
    assert all(sha(ROOT / name) == digest for name, digest in origin["unrelated"].items())
    origin.update(schema="litchi.performance.0824.origin.v1", base=base,
                  source_allowlist=["crates/litchi-pptx/src/opened/transaction.rs",
                                    "crates/litchi-pptx/src/opened/xml.rs"],
                  target=str(ROOT.parent / "litchi-target-0824"))
    write(P / "origin.json", origin)
    for family in ("synthetic", "real"):
        source = OLD / f"{family}-probe-src"
        destination = P / source.name
        shutil.copytree(source, destination)
        for path in destination.rglob("*"):
            if path.is_file() and path.suffix != ".lock":
                text = path.read_text().replace("change-0823", "change-0824")
                text = text.replace("litchi.performance.0823", "litchi.performance.0824")
                text = text.replace("b76786208d", base[:10])
                path.write_text(text)
        assert sha(destination / "Cargo.lock") == sha(source / "Cargo.lock")
    shutil.copytree(OLD / "inputs", P / "inputs")
    for name in ("root-inputs.json", "lock-parity.json", "corpus-inputs.json"):
        value = json.loads((OLD / name).read_text().replace("0823", "0824"))
        write(P / name, value)
    host = json.loads((OLD / "host.json").read_text())
    host.update(schema="litchi.performance.0824.host.v1",
                uname=subprocess.check_output(["uname", "-a"]).decode().strip(),
                target=str(ROOT.parent / "litchi-target-0824"),
                affinity_available=sorted(os.sched_getaffinity(0)),
                logical_cpu_count=os.cpu_count())
    host["mem_total"] = next(line for line in Path("/proc/meminfo").read_text().splitlines()
                             if line.startswith("MemTotal:"))
    host["mount"] = subprocess.check_output(
        ["findmnt", "-n", "-o", "SOURCE,FSTYPE,OPTIONS", "-T", str(ROOT)]).decode().strip()
    host["cgroup"] = {str(path): path.read_text().strip() for name in (
        "cpu.max", "memory.max", "cpuset.cpus.effective")
        if (path := Path("/sys/fs/cgroup") / name).is_file()}
    subprocess.run(["taskset", "-c", "12", "true"], check=True)
    write(P / "host.json", host)
    toolchain = json.loads((OLD / "toolchain.json").read_text())
    toolchain["schema"] = "litchi.performance.0824.toolchain.v1"
    for name in ("cargo -V", "rustc -Vv", "python3 --version", "git --version", "perf --version"):
        toolchain[name] = subprocess.check_output(name.split()).decode().strip()
    write(P / "toolchain.json", toolchain)
    print("0824 fresh input preparation PASS: unchanged architecture, corpus, locks, and unrelated files")


if __name__ == "__main__":
    main()
