"""Additive repair of the frozen driver's observer-identity admission typo.

No workload, binary, feature, sample, order or statistical rule changes. Existing
complete before qualifications are revalidated in place, never overwritten.
"""
import sys
import time
import driver as d

IDENTITY = "ordinary_save_procfs_and_system_allocator_operation_scoped"
RECOVERY = d.P / "recovery-freeze.json"
EVENTS = []


def freeze():
    d.check("before")
    assert d.read(d.P / "quality-before.json")["status"] == "pass"
    assert d.read(d.P / "build-before.json")["status"] == "pass"
    assert not (d.P / "qualification-before.json").exists()
    console = (d.P / "before-console.log").read_text()
    assert console.rstrip().endswith("AssertionError")
    assert 'assert report["tool"]["instrumentation"] ==' in console
    retained = [p for folder in ("commands", "qualification") for p in (d.P / folder).glob("*") if p.is_file()]
    retained += [d.P / name for name in ("driver.py", "recovery.py", "plan.json", "inputs.json", "prepare.json", "host.json", "before-console.log", "quality-before.json", "build-before.json")]
    d.write(RECOVERY, {"status": "frozen-before-resume", "base": d.BASE,
        "reason": "Observer features enable the exact procfs-decorated allocator identity; original driver expects allocator-only identity.",
        "permitted_change": "report identity admission only; no workload or analysis protocol change",
        "observer_identity": IDENTITY, "frozen_unix": time.time(),
        "files": {str(p.relative_to(d.P)): d.sha(p) for p in retained}})


def check():
    frozen = d.read(RECOVERY)
    assert frozen["status"] == "frozen-before-resume" and frozen["base"] == d.BASE
    assert frozen["observer_identity"] == IDENTITY
    assert all(d.sha(d.P / p) == h for p, h in frozen["files"].items())


def run_case(lane, block, leg, logical, case, samples, warmup, source_leg):
    check()
    d.check(source_leg)
    build = d.read(d.P / ("build-" + leg + ".json"))
    assert build["status"] == "pass"
    assert build["source_sha256"] == d.sha(d.P / "candidate" / (leg + "-validation.rs"))
    binding = build["binaries"][logical]
    assert binding["path"] == str(d.TARGET / (leg + "-" + logical))
    assert not d.Path(binding["path"]).is_symlink()
    assert d.sha(binding["path"]) == binding["sha256"]
    assert d.Path(binding["path"]).stat().st_size == binding["bytes"]
    stem = (leg + "-" + logical + "-" + case["id"]) if lane == "qualification" else f"{block:02}-{leg}-{case['id']}"
    dest = d.P / lane / (stem + ".json")
    dest.parent.mkdir(exist_ok=True)
    argv = ["taskset", "-c", "12", "/usr/bin/time", "-f", "%M", "-o", dest.with_suffix(".rss"),
            binding["path"], "--case", case["case"], "--warmup", str(warmup), "--samples", str(samples),
            "--json", dest, "--filesystem-root", d.SCRATCH]
    argv += ["--ooxml-file", d.INPUT] if case["shape"] is None else ["--xlsx-shape", case["shape"]]
    name = lane + "-" + stem
    terminal = d.P / "commands" / (name + ".json")
    start = d.P / "commands" / (name + ".started.json")
    log = d.P / "commands" / (name + ".log")
    paths = (dest, dest.with_suffix(".rss"), terminal, start, log)
    assert all(not p.is_symlink() for p in paths)
    present = [p.exists() for p in paths]
    reused = any(present)
    if reused:
        assert all(present), "partial artifact set cannot be resumed"
        frozen = d.read(RECOVERY)["files"]
        assert lane == "qualification" and leg == "before"
        assert all(str(p.relative_to(d.P)) in frozen for p in paths), "only prefrozen complete qualifications can be reused"
    else:
        d.run(name, argv, "capture", source_leg)
    receipt = d.read(terminal)
    started = d.read(start)
    assert all(receipt[k] == v for k, v in started.items())
    assert receipt["argv"] == list(map(str, argv)) and receipt["cwd"] == str(d.ROOT)
    assert receipt["exit_code"] == 0 and receipt["finished_unix"] >= receipt["started_unix"]
    assert receipt["source_leg"] == source_leg and receipt["source_sha256"] == d.sha(d.P / "candidate" / (source_leg + "-validation.rs"))
    assert receipt["input_inventory_sha256"] == d.sha(d.P / "inputs.json")
    assert receipt["log_sha256"] == d.sha(log)
    report = d.read(dest)
    assert report["configuration"]["samples_per_case"] == samples
    assert report["configuration"]["warmup_iterations_per_case"] == warmup
    assert report["configuration"]["cases"] == [case["case"]]
    assert len(report["results"]) == 1
    result = report["results"][0]
    assert result["case"] == case["case"] and len(result["elapsed_ns"]["samples"]) == samples
    assert report["binary_identity"]["binary_sha256"] == binding["sha256"]
    assert report["binary_identity"]["binary_bytes"] == binding["bytes"]
    assert result["corpus"]["shape"] == ("real-file" if case["shape"] is None else case["shape"])
    if case["shape"] is None:
        assert result["corpus"]["archive_sha256"] == d.sha(d.INPUT)
    else:
        assert case["shape"] in report["configuration"]["xlsx_shapes"]
    assert report["tool"]["instrumentation"] == ("none" if logical == "native" else IDENTITY)
    rss = dest.with_suffix(".rss").read_text().strip()
    assert rss.isdigit() and int(rss) > 0
    check()
    d.check(source_leg)
    EVENTS.append({"name": name, "reused": reused, "report_sha256": d.sha(dest), "command_sha256": d.sha(terminal)})
    print("revalidated" if reused else "admitted", name, flush=True)


def main():
    stage = sys.argv[1]
    if stage == "freeze":
        freeze()
        print("recovery frozen", flush=True)
        return
    assert stage in ("before", "after", "capture")
    check()
    d.run_case = run_case
    if stage == "before":
        d.check("before")
        d.qualification("before")
    elif stage == "after":
        d.check("after")
        d.quality("after")
        d.build("after")
        d.qualification("after")
    else:
        d.check("after")
        d.capture()
    check()
    d.write(d.P / ("recovery-" + stage + ".json"), {"status": "pass", "stage": stage,
        "recovery_freeze_sha256": d.sha(RECOVERY), "events": EVENTS})
    print("recovery", stage, "PASS", flush=True)


if __name__ == "__main__":
    main()
