"""Root preparation of fresh 0828 protocol, host, and immutable input copies."""
import hashlib
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import time
import tomllib

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / "change-0822"
BASE = "c990602492106d968898b7310bdd4bc17f3e4fcb"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path):
    return json.loads(path.read_text())


def write(name, value):
    path = P / name
    assert not path.exists(), path
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def command(argv):
    return subprocess.check_output(argv, cwd=ROOT, text=True)


def main():
    assert command(["git", "rev-parse", "HEAD"]).strip() == BASE
    assert not (ROOT.parent / "litchi-target-0828").exists()
    architecture = read(P.parent / "change-0827/architecture-inputs.json")
    assert len(architecture) == 35
    assert all(sha(ROOT / name) == digest for name, digest in architecture.items())
    write("architecture-inputs.json", architecture)
    origin = json.loads((OLD / "origin.json").read_text().replace("0822", "0828"))
    assert all(sha(ROOT / name) == digest for name, digest in origin["unrelated"].items())
    origin.update(base=BASE, previous_commit=BASE,
                  previous_packet="docs/performance/results/change-0827")
    origin["quality_reuse"] = {
        "schema": "litchi.performance.0828.quality-reuse.v1",
        "mode": "exact-source-0827-quality-receipts",
        "cargo_commands_executed": False,
        "prior_quality": "docs/performance/results/change-0827/quality.json",
        "prior_seal": "docs/performance/results/change-0827/seal.json",
        "prior_seal_commit": BASE,
        "required_test_counts": {"passed": 641, "failed": 0, "ignored": 1, "suites": 28},
    }
    write("origin.json", origin)
    plan = json.loads((OLD / "plan.json").read_text().replace("0822", "0828"))
    plan["base"] = BASE
    plan["purpose"] = "Fresh sampled phase ownership for the current real-file PPTX public edit, with direct/wrapped/frame-pointer controls."
    plan["statistics"]["bootstrap"]["seed"] = 828828
    plan["quality"]["production_mode"] = "exact-source-0827-quality-receipts"
    phase_owners = {key: "pptx_edit_profile_0828::" + name for key, name in (
        ("capture", "phase_opened_presentation_transaction_0828"),
        ("set_text", "phase_set_shape_text_0828"),
        ("publish", "phase_commit_apply_opened_presentation_commit_0828"),
    )}
    plan["probe"]["phase_owners"] = phase_owners
    plan["perf"]["phase_owners"] = phase_owners
    write("plan.json", plan)
    for name in ("corpus-inputs.json", "input-inventory.json", "provenance.json"):
        value = json.loads((OLD / name).read_text().replace("0822", "0828"))
        if name == "input-inventory.json":
            value["scope"] = "Open the checked-in shapes.pptx input; serialize edited bytes in memory and compare against the pinned 0821 default output. No filesystem scratch or atomic publication timing."
        if name == "provenance.json":
            value["quality_reuse"] = "docs/performance/results/change-0827/quality.json"
        write(name, value)
    (P / "inputs").mkdir()
    for source, destination in (("Cargo.lock", "root-Cargo.lock"),
                                ("rustfmt.toml", "rustfmt.toml"),
                                ("tools/perf-baseline/Cargo.lock", "tool-Cargo.lock")):
        shutil.copy2(ROOT / source, P / "inputs" / destination)
    write("root-inputs.json", {name: sha(ROOT / name) for name in ("Cargo.lock", "rustfmt.toml")})
    roots = tomllib.loads((ROOT / "Cargo.lock").read_text())["package"]
    tool = tomllib.loads((ROOT / "tools/perf-baseline/Cargo.lock").read_text())["package"]
    def identity(row):
        return tuple(row.get(k) for k in ("name", "version", "source", "checksum"))
    external = [identity(row) for row in tool if "source" in row]
    root_set = {identity(row) for row in roots}
    missing = sorted(row for row in external if row not in root_set)
    write("lock-parity.json", {"schema": "litchi.performance.0828.lock-parity.v1",
                              "root_packages": len(roots), "tool_packages": len(tool),
                              "external_tool_packages": len(external), "missing": missing,
                              "all_external_tool_entries_match_root": not missing})
    host = {"observed": time.time(), "platform": platform.platform(),
            "affinity_available": sorted(os.sched_getaffinity(0)), "affinity_selected": [12],
            "filesystem": command(["findmnt", "-n", "-o", "SOURCE,FSTYPE,OPTIONS", "-T", str(ROOT)]),
            "filesystem_capture_parent": str(ROOT.parent)}
    for field in ("cpuinfo", "meminfo", "loadavg"):
        host[field] = (Path("/proc") / field).read_text()
    host["cgroup"] = Path("/proc/self/cgroup").read_text()
    subprocess.run(["taskset", "-c", "12", "true"], check=True)
    write("host.json", host)
    write("toolchain.json", {" ".join(argv): command(argv) for argv in
                             (["rustc", "-Vv"], ["cargo", "-V"], ["python3", "--version"],
                              ["git", "--version"], ["perf", "--version"])})
    current = Path("/sys/fs/cgroup") / host["cgroup"].split("0::", 1)[1].strip().lstrip("/")
    rows = []
    while current.is_relative_to(Path("/sys/fs/cgroup")):
        rows.append({"path": str(current), **{name: (current / name).read_text().strip()
                     if (current / name).is_file() else None
                     for name in ("cpu.max", "memory.max", "cpuset.cpus.effective")}})
        if current == Path("/sys/fs/cgroup"):
            break
        current = current.parent
    write("cgroup-limits.json", rows)
    print("0828 preparation PASS: current base, unchanged normative and unrelated inputs, fresh host and lock census")


if __name__ == "__main__":
    main()
