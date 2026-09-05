#!/usr/bin/env python3
"""Fail-closed resource evidence checker for goal 0413.

This only reads capture artifacts.  It does not build, run a benchmark, or
invoke perf.  The primary goal validator owns schema/corpus/sample checks;
this checker binds the command journal to the protocol/build identities and
reports the captured RSS and PMU evidence.
"""

import argparse
import hashlib
import json
import math
import re
import statistics
import subprocess
import sys
from pathlib import Path


EVENTS = [
    "task-clock", "cycles", "instructions", "branches", "branch-misses",
    "page-faults", "context-switches", "cpu-migrations",
    "l2_cache_req_stat.dc_access_in_l2", "l2_cache_req_stat.dc_hit_in_l2",
]
SOFTWARE_EVENTS = {"task-clock", "page-faults", "context-switches", "cpu-migrations"}
REVISION = {
    "control": "ceba0345220c1ca6a7f61f3fac86145b5afc55ca",
    "candidate": "bf5b7f50f61ba17091ef80dc509b64378b11aaa7",
}


class Problem(Exception):
    pass


def need(condition, message):
    if not condition:
        raise Problem(message)


def payload(path):
    """Read path, falling back to path + .zst for published captures."""
    path = Path(path)
    if path.exists():
        return path.read_bytes()
    compressed = Path(str(path) + ".zst")
    need(compressed.exists(), f"missing artifact: {path} (or {compressed})")
    try:
        result = subprocess.run(
            ["zstd", "-q", "-dc", "--", str(compressed)],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as exc:
        raise Problem(f"cannot decompress {compressed}: {exc}") from exc
    return result.stdout


def read_json(path):
    try:
        return json.loads(payload(path))
    except (ValueError, UnicodeDecodeError) as exc:
        raise Problem(f"invalid JSON: {path}: {exc}") from exc


def read_text(path):
    try:
        return payload(path).decode()
    except UnicodeDecodeError as exc:
        raise Problem(f"invalid UTF-8 text: {path}: {exc}") from exc


def digest(path):
    h = hashlib.sha256()
    with open(path, "rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def first_line(value):
    return str(value).splitlines()[0] if str(value).splitlines() else ""


def artifact(root, label, suffix):
    return Path(root) / f"{label}.{suffix}"


def classify(label):
    if label in ("control-profile", "candidate-profile"):
        return "profile", None
    for prefix, kind in (
        ("normal-", "normal"), ("allocator-", "allocator"),
        ("guard-normal-", "guard-normal"), ("guard-allocator-", "guard-allocator"),
        ("stat-", "stat"),
    ):
        if label.startswith(prefix):
            try:
                index = int(label[len(prefix):])
            except ValueError:
                break
            return kind, index
    raise Problem(f"unknown command label: {label}")


def expected_labels(labels):
    main = (
        [f"normal-{i}" for i in range(1, 5)]
        + [f"allocator-{i}" for i in range(1, 5)]
        + [f"guard-normal-{i}" for i in range(1, 5)]
        + [f"guard-allocator-{i}" for i in range(1, 5)]
        + ["control-profile", "candidate-profile"]
        + [f"stat-{i}" for i in range(1, 5)]
    )
    recheck = [f"guard-normal-{i}" for i in range(1, 5)]
    actual = list(labels)
    if actual == main:
        return main, False
    if actual == recheck:
        return recheck, True
    raise Problem(f"unexpected command label/order: {actual}")


def identity_paths(root):
    if root not in (Path("/tmp/litchi-goal-0413-capture"), Path("/tmp/litchi-goal-0413-guard-recheck")):
        return {role: root.parent / "checks" / (role+"-build-identity.json") for role in ("control","candidate")}
    return {
        "control": Path("/tmp/litchi-goal-0413-control-binaries/identity.json"),
        "candidate": Path("/tmp/litchi-goal-0413-candidate-binaries/identity.json"),
    }


def expected_variant(proto, kind, index):
    if kind in ("normal", "allocator", "guard-normal", "guard-allocator"):
        section = proto["guards"] if kind.startswith("guard-") else proto
        key = "allocator" if kind.endswith("allocator") else "normal"
        return section[key]["order"][index - 1]
    if kind == "stat":
        roles = proto["counters"]["roles"]
        return [roles[0], roles[1], roles[1], roles[0]][index - 1]
    if kind == "profile":
        return "control" if index is None else None
    raise Problem(f"no variant rule for {kind}")


def executable_for(root, identities, variant, allocator):
    name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    path = Path(f"/tmp/litchi-goal-0413-{variant}-binaries") / name
    return str(path)


def common_argv(root, label, binary, case, samples, warmup, proto):
    return [
        binary, "--filesystem-cache", proto["filesystem_cache"], "--case", case,
        "--samples", str(samples), "--warmup", str(warmup),
        "--json", str(artifact(root, label, "json")),
        "--corpus-manifest", str(artifact(root, label, "catalog.json")),
    ]


def expected_argv(root, label, kind, index, variant, proto, identities):
    guard = kind.startswith("guard-")
    allocator = kind in ("allocator", "guard-allocator")
    binary = executable_for(root, identities, variant, allocator)
    if kind in ("normal", "allocator", "guard-normal", "guard-allocator"):
        section = proto["guards"] if guard else proto
        mode = "allocator" if allocator else "normal"
        cfg = section[mode]
        case = ",".join(section["cases"] if guard else proto["cases"])
        argv = ["taskset", "-c", str(proto["cpu"]), "/usr/bin/time", "-v", "-o",
                str(artifact(root, label, "time.txt"))]
        argv += common_argv(root, label, binary, case, cfg["samples"], cfg["warmup"], proto)
        if guard:
            argv += ["--shape", ",".join(section["shapes"]),
                     "--payload", ",".join(section["payloads"])]
        return argv
    if kind == "profile":
        cfg = proto["profile"]
        return [
            "taskset", "-c", str(proto["cpu"]), "perf", "record",
            "--no-buildid-cache", "-e", cfg["event"], "-F", str(cfg["frequency"]),
            "--call-graph", cfg["call_graph"], "-o", str(artifact(root, label, "data")), "--",
        ] + common_argv(root, label, binary, cfg["case"], cfg["samples"], cfg["warmup"], proto)
    if kind == "stat":
        cfg = proto["counters"]
        return [
            "taskset", "-c", str(proto["cpu"]), "perf", "stat", "--no-big-num", "-x,",
            "-e", cfg["events"], "-o", str(artifact(root, label, "csv")), "--",
        ] + common_argv(root, label, binary, cfg["case"], cfg["samples"], cfg["warmup"], proto)
    raise Problem(f"no argv rule for {kind}")


def validate_identity_files(identities):
    for variant, ident in identities.items():
        need(ident["revision"] == REVISION[variant], f"{variant}: revision mismatch")
        need(ident.get("source_status", "") == "", f"{variant}: source is not clean")
        env = ident.get("build_environment", {})
        need(env.get("RUSTUP_TOOLCHAIN") == "1.98.1", f"{variant}: toolchain mismatch")
        need(env.get("RUSTFLAGS") == "-C force-frame-pointers=yes -C force-unwind-tables=yes",
             f"{variant}: RUSTFLAGS mismatch")
        need(set(ident.get("binaries", {})) == {"litchi-perf-baseline","litchi-perf-baseline-alloc"}, f"{variant}: binary set mismatch")
        for name, info in ident["binaries"].items():
            need(isinstance(info["bytes"],int) and info["bytes"]>0 and re.fullmatch(r"[0-9a-f]{64}",info["sha256"]), f"{variant}: malformed binary metadata")
            if not ident["_live"]:
                continue  # Published replay binds build sidecar, journal, and report; capture hashed actual binaries.
            path = Path(ident["_path"]).parent / name
            need(path.is_file(), f"{variant}: missing binary {path}")
            need(path.stat().st_size == info["bytes"], f"{variant}: {name} byte size mismatch")
            need(digest(path) == info["sha256"], f"{variant}: {name} sha256 mismatch")


def validate_report(root, label, kind, variant, proto, identities, journal):
    report = read_json(artifact(root, label, "json"))
    need(isinstance(report, dict), f"{label}: report is not an object")
    ident = identities[variant]
    allocator = kind in ("allocator", "guard-allocator")
    binary_name = "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline"
    expected_binary = Path(f"/tmp/litchi-goal-0413-{variant}-binaries") / binary_name
    tool = report.get("tool", {})
    need(report.get("schema_version") == 1, f"{label}: schema version mismatch")
    need(tool.get("binary") == binary_name, f"{label}: tool binary mismatch")
    need(tool.get("instrumentation") == ("system_allocator_operation_scoped" if allocator else "none"),
         f"{label}: instrumentation mismatch")
    bi = report.get("binary_identity", {})
    info = ident["binaries"][binary_name]
    need(bi.get("path") == str(expected_binary), f"{label}: report binary path mismatch")
    need(bi.get("binary_sha256") == info["sha256"] == journal["binary_sha256"],
         f"{label}: report/binary identity hash mismatch")
    need(bi.get("binary_bytes") == info["bytes"], f"{label}: report binary size mismatch")
    need(bi.get("executable") is True and bi.get("profile") == "release",
         f"{label}: binary executable/profile mismatch")
    env = report.get("environment", {})
    need(env.get("git_revision") == ident["revision"], f"{label}: report revision mismatch")
    need(env.get("git_worktree_dirty") is False, f"{label}: dirty report worktree")
    need(env.get("rustc_version") == first_line(ident["rustc"]), f"{label}: rustc mismatch")
    need(env.get("rustflags") == ident["build_environment"]["RUSTFLAGS"], f"{label}: rustflags mismatch")
    need(env.get("cpu_affinity") == str(proto["cpu"]), f"{label}: CPU affinity mismatch")
    need(env.get("os") == "linux" and env.get("logical_cpus_available") == 1,
         f"{label}: environment/worker identity mismatch")
    ident_kernel = ident.get("kernel", "")
    need(env.get("kernel") and env["kernel"].split()[-1] in ident_kernel,
         f"{label}: kernel identity mismatch")
    need(env.get("cpu_model") and env["cpu_model"] in ident.get("cpu", ""),
         f"{label}: CPU model identity mismatch")
    cfg = report.get("configuration", {})
    if kind == "profile":
        expected = proto["profile"]
        cases = [expected["case"]]
    elif kind == "stat":
        expected = proto["counters"]
        cases = [expected["case"]]
    else:
        section = proto["guards"] if kind.startswith("guard-") else proto
        mode = "allocator" if kind.endswith("allocator") else "normal"
        expected = section[mode]
        cases = section["cases"] if kind.startswith("guard-") else proto["cases"]
    need(cfg.get("samples_per_case") == expected["samples"], f"{label}: sample count mismatch")
    need(cfg.get("warmup_iterations_per_case") == expected["warmup"], f"{label}: warmup mismatch")
    need(cfg.get("cases") == cases, f"{label}: input case identity mismatch")
    need(cfg.get("filesystem_cache_states") == [proto["filesystem_cache"]], f"{label}: cache mode mismatch")
    need(cfg.get("filesystem_fresh_child_per_sample") is True and cfg.get("filesystem_process_isolated") is True,
         f"{label}: isolation identity mismatch")
    need(isinstance(report.get("corpus_catalog"), dict) and isinstance(report.get("results"), list),
         f"{label}: report corpus/results missing")
    catalog = read_json(artifact(root, label, "catalog.json"))
    for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
        need(report["corpus_catalog"].get(key) == catalog.get(key),
             f"{label}: report/catalog identity mismatch for {key}")
    need(catalog.get("build", {}).get("git_revision") == ident["revision"] and
         catalog.get("build", {}).get("git_worktree_dirty") is False,
         f"{label}: catalog build identity mismatch")
    if kind.startswith("guard-"):
        need(cfg.get("corpus_shapes") == proto["guards"]["shapes"], f"{label}: guard shape identity mismatch")
        need(cfg.get("payload_kinds") == proto["guards"]["payloads"], f"{label}: guard payload identity mismatch")
    return report


def parse_rss(root, label):
    text = read_text(artifact(root, label, "time.txt"))
    m = re.search(r"Maximum resident set size \(kbytes\):\s*([0-9]+)", text)
    need(m, f"{label}: missing maximum RSS")
    rss = int(m.group(1))
    need(rss > 0, f"{label}: non-positive RSS")
    status = re.search(r"Exit status:\s*(-?\d+)", text)
    need(status and int(status.group(1)) == 0, f"{label}: timed command failed")
    elapsed = re.search(r"Elapsed \(wall clock\) time \([^)]*\):\s*([^\n]+)", text)
    need(elapsed and elapsed.group(1).strip() not in ("0", "0:00", "0:00.00"),
         f"{label}: non-positive elapsed time")
    return rss


def rss_group(root, labels, variants, name):
    values = {label: parse_rss(root, label) for label in labels}
    controls = [values[l] for l, v in zip(labels, variants) if v == "control"]
    candidates = [values[l] for l, v in zip(labels, variants) if v == "candidate"]
    need(controls and candidates, f"{name}: incomplete control/candidate RSS pair")
    pair_deltas = [100.0 * abs(c - b) / b for b in controls for c in candidates]
    median_delta = 100.0 * abs(statistics.median(candidates) - statistics.median(controls)) / statistics.median(controls)
    need(max(pair_deltas) <= 5.0, f"{name}: RSS pair exceeds 5% ({max(pair_deltas):.3f}%)")
    return {
        "labels": labels, "variants": variants, "rss_kib": values,
        "control_values_kib": controls, "candidate_values_kib": candidates,
        "control_median_kib": statistics.median(controls),
        "candidate_median_kib": statistics.median(candidates),
        "median_delta_percent": median_delta,
        "maximum_pair_delta_percent": max(pair_deltas),
        "limit_percent": 5.0,
    }


def parse_stat(root, label, proto):
    text = read_text(artifact(root, label, "csv"))
    rows = []
    for line in text.splitlines():
        if not line.strip() or line.startswith("#"):
            continue
        fields = line.split(",")
        need(len(fields) >= 5, f"{label}: malformed perf stat row")
        event = fields[2]
        need(event not in {r["event"] for r in rows}, f"{label}: duplicate PMU event {event}")
        try:
            count = float(fields[0].replace(",", ""))
            running_ns = float(fields[3].replace(",", ""))
            scheduled = float(fields[4].replace(",", ""))
        except ValueError as exc:
            raise Problem(f"{label}: non-numeric PMU row {line!r}") from exc
        need(all(math.isfinite(v) for v in (count, running_ns, scheduled)), f"{label}: non-finite PMU value")
        need(running_ns > 0, f"{label}/{event}: non-positive runtime")
        need(scheduled >= 80.0, f"{label}/{event}: scheduling below 80%")
        if event in SOFTWARE_EVENTS:
            need(abs(scheduled - 100.0) < 1e-6, f"{label}/{event}: software scheduling is not 100%")
        need(count >= 0 if event == "cpu-migrations" else count > 0,
             f"{label}/{event}: non-positive scaled count")
        rows.append({
            "event": event, "class": "software" if event in SOFTWARE_EVENTS else "hardware",
            "scaled_count": count, "unit": fields[1],
            "runtime_ns": running_ns, "scheduled_percent": scheduled,
        })
    need([r["event"] for r in rows] == EVENTS, f"{label}: PMU event set/order mismatch")
    values = {r["event"]: r["scaled_count"] for r in rows}
    cycles, instructions = values["cycles"], values["instructions"]
    need(cycles > 0, f"{label}: no cycles for IPC")
    ipc = instructions / cycles
    return {
        "scope": proto["counters"]["scope"], "claim": "descriptive",
        "events": rows, "whole_process_scaled_ipc": ipc,
    }


def profile_evidence(root, label, proto):
    data = payload(artifact(root, label, "data"))
    need(data, f"{label}: empty perf.data")
    text = read_text(artifact(root, label, "self.txt"))
    lost = re.search(r"Total Lost Samples:\s*([0-9]+)", text)
    event = re.search(r"event '([^']+)'", text)
    count = re.search(r"Event count \(approx\.\):\s*([0-9,]+)", text)
    need(lost and int(lost.group(1)) == 0, f"{label}: lost profile samples")
    need(event and event.group(1) == proto["profile"]["event"], f"{label}: profile event mismatch")
    need(count and int(count.group(1).replace(",", "")) > 0, f"{label}: empty profile event count")
    for suffix in ("header.txt", "script.txt"):
        need(payload(artifact(root, label, suffix)), f"{label}: missing profile {suffix}")
    return {
        "event": event.group(1), "lost_samples": int(lost.group(1)),
        "event_count": int(count.group(1).replace(",", "")), "data_bytes": len(data),
    }


def check_capture(root):
    root = Path(root).resolve()
    protocol_path = root / "protocol.json"
    proto_bytes = payload(protocol_path)
    proto = json.loads(proto_bytes)
    need(proto.get("change") == "0413", f"{root}: wrong protocol change")
    protocol_sha = hashlib.sha256(proto_bytes).hexdigest()
    paths = identity_paths(root)
    identities = {}
    for variant, path in paths.items():
        ident = read_json(path)
        ident["_path"] = str(path)
        ident["_live"] = path.parent == Path(f"/tmp/litchi-goal-0413-{variant}-binaries")
        identities[variant] = ident
    validate_identity_files(identities)
    journal = read_json(root / "commands.json")
    need(isinstance(journal, list) and journal, f"{root}: empty command journal")
    labels = [item.get("label") for item in journal]
    expected, recheck = expected_labels(labels)
    records = []
    reports = {}
    rss = {}
    pmu = {}
    profiles = {}
    for item, label in zip(journal, expected):
        kind, index = classify(label)
        if kind == "profile":
            variant = "control" if label == "control-profile" else "candidate"
        else:
            variant = expected_variant(proto, kind, index)
        need(item.get("label") == label and item.get("variant") == variant,
             f"{label}: journal order/variant mismatch")
        ident = identities[variant]
        need(item.get("revision") == ident["revision"], f"{label}: journal revision mismatch")
        need(item.get("source_status", "") == ident.get("source_status", ""), f"{label}: source status mismatch")
        need(item.get("protocol_sha256") == protocol_sha, f"{label}: protocol sha mismatch")
        need(item.get("cwd") == ident["build_cwd"], f"{label}: build cwd mismatch")
        need(item.get("exit_code") == 0 and float(item.get("wall_seconds", 0)) > 0,
             f"{label}: command did not complete with positive wall time")
        argv = expected_argv(Path("/tmp/litchi-goal-0413-guard-recheck" if recheck else "/tmp/litchi-goal-0413-capture"), label, kind, index, variant, proto, identities)
        need(item.get("argv") == argv, f"{label}: argv/input identity mismatch")
        expected_binary = executable_for(root, identities, variant, kind in ("allocator", "guard-allocator"))
        binary_name = Path(expected_binary).name
        need(item.get("binary_sha256") == ident["binaries"][binary_name]["sha256"],
             f"{label}: journal binary identity mismatch")
        reports[label] = validate_report(root, label, kind, variant, proto, identities, item)
        record = {k: item[k] for k in ("label", "variant", "revision", "protocol_sha256", "binary_sha256", "argv", "cwd")}
        records.append(record)
        if kind == "stat":
            pmu[label] = parse_stat(root, label, proto)
        if kind == "profile":
            profiles[label] = profile_evidence(root, label, proto)

    if not recheck:
        for prefix, name in (("normal", "normal"), ("allocator", "allocator"),
                             ("guard-normal", "guard-normal"), ("guard-allocator", "guard-allocator")):
            labels4 = [f"{prefix}-{i}" for i in range(1, 5)]
            variants4 = [next(x["variant"] for x in journal if x["label"] == label) for label in labels4]
            rss[name] = rss_group(root, labels4, variants4, name)
    else:
        labels4 = [f"guard-normal-{i}" for i in range(1, 5)]
        variants4 = [next(x["variant"] for x in journal if x["label"] == label) for label in labels4]
        rss["guard-normal-recheck"] = rss_group(root, labels4, variants4, "guard-normal-recheck")

    return {
        "root": str(root), "protocol": {
            "path": str(protocol_path), "sha256": protocol_sha,
            "change": proto["change"], "control_revision": proto["control_revision"],
            "candidate_revision": proto["candidate_revision"],
        },
        "recheck": recheck, "identities": {
            v: {"path": i["_path"], "revision": i["revision"], "source_status": i.get("source_status", ""),
                "binaries": i["binaries"]} for v, i in identities.items()
        },
        "commands": records, "rss": rss, "pmu": pmu, "profiles": profiles,
        "report_count": len(reports),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", help="capture root (default: /tmp/litchi-goal-0413-capture)")
    parser.add_argument("--root", dest="root_option")
    parser.add_argument("--recheck", default="/tmp/litchi-goal-0413-guard-recheck")
    parser.add_argument("--output", default="/tmp/litchi-goal-0413-resources.json")
    args = parser.parse_args()
    root = args.root_option or args.root or "/tmp/litchi-goal-0413-capture"
    result = {"schema_version": 1, "change": "0413", "status": "fail", "errors": []}
    try:
        result["capture"] = check_capture(root)
        recheck = Path(args.recheck)
        if recheck.exists():
            result["guard_recheck"] = check_capture(recheck)
        else:
            raise Problem(f"missing required guard recheck root: {recheck}")
        result["status"] = "pass"
    except (Problem, OSError, KeyError, TypeError, ValueError) as exc:
        result["errors"].append(str(exc))
    output = Path(args.output)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": result["status"], "output": str(output), "errors": result["errors"]}))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    sys.exit(main())
