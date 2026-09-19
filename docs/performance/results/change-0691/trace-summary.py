#!/usr/bin/env python3
"""Verify and summarize completed 0691 temporary trace runs."""
from __future__ import annotations
import argparse, hashlib, json, re, sys
from collections import Counter, defaultdict
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
RUNS = P / "trace-runs"
VERIFY_LIVE_BINARY = False
HEX = re.compile(r"^0x[0-9a-fA-F]+$")

def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()
def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))
def dump(path: Path, value) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
def integer(value: str) -> int:
    return int(value, 0) if HEX.fullmatch(value) else int(value)
def fields(line: str):
    words = line.strip().split()
    if not words or not words[0].startswith("LITCHI0691_"):
        return None, {}
    result = {}
    for word in words[1:]:
        key, sep, value = word.partition("=")
        if sep:
            result[key] = value
    return words[0], result

def verify_files(run: Path, manifest: dict) -> list[str]:
    errors = []
    if manifest.get("status") != "completed":
        return ["manifest status is not completed"]
    if manifest.get("source_restored_exact") is not True:
        errors.append("manifest source_restored_exact is not true")
    restoration = manifest.get("restoration", [])
    if not restoration:
        errors.append("manifest has no restoration records")
    for item in restoration:
        if item.get("restored_exact") is not True:
            errors.append(f"restoration failed for {item.get('path')}")
        path = ROOT / item["path"]
        if not path.is_file() or sha(path) != item.get("restored_sha256"):
            errors.append(f"live source hash differs after restore: {item.get('path')}")
    hashes_path = run / "source-hashes.json"
    patch_path = run / "trace.patch"
    if not hashes_path.is_file() or not patch_path.is_file():
        return errors + ["source hash archive or trace.patch is missing"]
    hashes = load(hashes_path)
    if sha(patch_path) != hashes.get("patch_sha256"):
        errors.append("trace.patch hash differs from source-hashes.json")
    for item in hashes.get("files", []):
        stem = item["path"].replace("/", "__")
        before = run / "sources" / (stem + ".before")
        after = run / "sources" / (stem + ".after")
        if not before.is_file() or not after.is_file():
            errors.append(f"source archive missing for {item['path']}")
            continue
        if sha(before) != item.get("before_sha256"):
            errors.append(f"archived before hash differs for {item['path']}")
        if sha(after) != item.get("after_sha256"):
            errors.append(f"archived after hash differs for {item['path']}")
        restored = next((x for x in restoration if x.get("path") == item["path"]), None)
        if restored is None:
            errors.append(f"restoration record is missing for {item['path']}")
        elif restored.get("restored_sha256") != item.get("before_sha256"):
            errors.append(f"restore hash does not equal archived before for {item['path']}")
    records = load(run / "probe-runs.json") if (run / "probe-runs.json").is_file() else []
    if records != manifest.get("probe_runs", []):
        errors.append("probe-runs.json count differs from manifest")
    for record in records:
        if record.get("exit_code") != 0:
            errors.append(f"probe leg failed: {record.get('case')}/{record.get('operation')}")
        label = f"{record['case']}-{record['operation']}-{record['count']}"
        for suffix, field in ((".stdout", "stdout_sha256"), (".stderr", "stderr_sha256")):
            path = run / (label + suffix)
            if not path.is_file() or sha(path) != record.get(field):
                errors.append(f"{label}{suffix} hash differs from receipt")
    binary = manifest.get("binary")
    if VERIFY_LIVE_BINARY and binary and manifest.get("binary_sha256"):
        path = Path(binary)
        if not path.is_file() or sha(path) != manifest["binary_sha256"]:
            errors.append("trace binary hash differs from manifest")
    return errors

def event_interval(role: str, ordinal: int, total: int) -> str:
    if role == "apply":
        if ordinal == 0: return "derive_target_setup"
        if ordinal == total - 2: return "route_source"
        if ordinal == total - 1: return "route_commit"
        return "unexpected_intermediate"
    return "measured_capture"

def summarize_events(codec: list[dict], parts: list[dict]) -> dict:
    uri_by_key = defaultdict(set)
    alias_by_key = defaultdict(set)
    for part in parts:
        key = (part["visible_ptr"], part["visible_len"])
        uri_by_key[key].add(part["uri"])
        alias_by_key[key].add(part["alias"])
    pointers = {}
    modes = Counter()
    statuses = Counter()
    inputs = outputs = 0
    for call in codec:
        key = (call["input_ptr"], call["input_len"])
        raw_key = f"0x{key[0]:x}/{key[1]}"
        row = pointers.setdefault(raw_key, {
            "input_ptr": f"0x{key[0]:x}", "input_len": key[1], "calls": 0,
            "input_bytes": 0, "output_bytes": 0, "mode_counts": {},
            "status_counts": {}, "output_ptrs": [], "mapped_uris": [],
            "part_alias_values": [],
        })
        row["calls"] += 1
        row["input_bytes"] += call["input_len"]
        row["output_bytes"] += call["output_len"]
        row["mode_counts"][call["mode"]] = row["mode_counts"].get(call["mode"], 0) + 1
        row["status_counts"][call["status"]] = row["status_counts"].get(call["status"], 0) + 1
        output_ptr = f"0x{call['output_ptr']:x}"
        if output_ptr not in row["output_ptrs"]: row["output_ptrs"].append(output_ptr)
        row["mapped_uris"] = sorted(uri_by_key.get(key, ()))
        row["part_alias_values"] = sorted(alias_by_key.get(key, ()))
        modes[call["mode"]] += 1
        statuses[call["status"]] += 1
        inputs += call["input_len"]; outputs += call["output_len"]
    return {
        "codec_calls": len(codec), "input_bytes": inputs, "output_bytes": outputs,
        "borrowed_calls": modes.get("borrow", 0), "owned_calls": modes.get("owned", 0),
        "error_calls": statuses.get("error", 0), "mode_counts": dict(modes),
        "status_counts": dict(statuses), "part_records": len(parts),
        "raw_pointer_len": sorted(pointers.values(), key=lambda x: (x["input_ptr"], x["input_len"])),
        "pointer_scope": "one probe process; do not compare addresses across stderr files",
    }

def parse_trace(path: Path, operation: str) -> dict:
    intervals, outside = [], {"codec": [], "parts": []}
    all_events = {"codec": [], "parts": []}
    active = None
    errors = []
    for line_no, line in enumerate(path.read_text(errors="replace").splitlines(), 1):
        prefix, values = fields(line)
        if prefix == "LITCHI0691_MODEL":
            stage = values.get("stage")
            if stage == "capture_begin":
                if active is not None: errors.append(f"{path.name}:{line_no}: nested capture_begin")
                active = {"begin_line": line_no, "codec": [], "parts": []}
            elif stage == "capture_success":
                if active is None: errors.append(f"{path.name}:{line_no}: capture_success without begin")
                else:
                    active["end_line"] = line_no; intervals.append(active); active = None
            continue
        target = active if active is not None else outside
        if prefix == "LITCHI0691_CODEC":
            try:
                event = {
                    "line": line_no, "call": int(values["call"]),
                    "input_ptr": integer(values["input_ptr"]), "input_len": int(values["input_len"]),
                    "output_ptr": integer(values["output_ptr"]), "output_len": int(values["output_len"]),
                    "mode": values["mode"], "status": values["status"],
                }
                target["codec"].append(event); all_events["codec"].append(event)
            except (KeyError, ValueError) as exc:
                errors.append(f"{path.name}:{line_no}: malformed codec record: {exc}")
        elif prefix == "LITCHI0691_PART":
            try:
                event = {
                    "line": line_no, "uri": values["uri"],
                    "visible_ptr": integer(values["visible_ptr"]),
                    "visible_len": int(values["visible_len"]),
                    "arc_bytes_ptr": integer(values["arc_bytes_ptr"]),
                    "arc_len": int(values["arc_len"]), "alias": values["alias"],
                }
                target["parts"].append(event); all_events["parts"].append(event)
            except (KeyError, ValueError) as exc:
                errors.append(f"{path.name}:{line_no}: malformed part record: {exc}")
    if active is not None:
        errors.append(f"{path.name}: capture_begin has no capture_success")
        intervals.append(active)
    if not intervals:
        errors.append(f"{path.name}: no capture boundaries")
    total = len(intervals)
    summaries = []
    for ordinal, item in enumerate(intervals):
        summary = summarize_events(item["codec"], item["parts"])
        summary.update({"ordinal": ordinal, "role": event_interval(operation, ordinal, total),
                        "begin_line": item["begin_line"], "end_line": item.get("end_line"),
                        "call_first": item["codec"][0]["call"] if item["codec"] else None,
                        "call_last": item["codec"][-1]["call"] if item["codec"] else None})
        summaries.append(summary)
    outside_summary = summarize_events(outside["codec"], outside["parts"])
    return {"intervals": summaries, "outside_capture": outside_summary,
            "process": summarize_events(all_events["codec"], all_events["parts"]),
            "errors": errors}

def summarize_run(run: Path) -> dict:
    manifest = load(run / "manifest.json")
    errors = verify_files(run, manifest)
    legs = []
    for record in manifest.get("probe_runs", []):
        label = f"{record['case']}-{record['operation']}-{record['count']}"
        stderr = run / (label + ".stderr")
        parsed = parse_trace(stderr, record["operation"])
        errors.extend(f"{label}: {error}" for error in parsed["errors"])
        legs.append({"case": record["case"], "source": record["source"], "operation": record["operation"],
                     "count": record["count"], "stderr": stderr.name,
                     "intervals": parsed["intervals"], "outside_capture": parsed["outside_capture"],
                     "process": parsed["process"]})
    return {"run": str(run), "profile": manifest.get("profile"), "model_trace": manifest.get("model_trace"),
            "verification_errors": errors, "verification": "pass" if not errors else "fail", "legs": legs}

def main() -> int:
    global VERIFY_LIVE_BINARY
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_dirs", nargs="*", help="completed trace run directories")
    parser.add_argument("--output", default=str(P / "trace-summary.json"))
    parser.add_argument("--verify-live-binary", action="store_true", help="also require the temporary executable to still exist and match")
    args = parser.parse_args()
    VERIFY_LIVE_BINARY = args.verify_live_binary
    paths = [Path(x) for x in args.run_dirs] if args.run_dirs else sorted(x for x in RUNS.iterdir() if x.is_dir())
    reports, skipped = [], []
    for run in paths:
        manifest_path = run / "manifest.json"
        if not manifest_path.is_file():
            skipped.append(str(run)); continue
        manifest = load(manifest_path)
        if manifest.get("status") != "completed":
            if args.run_dirs: reports.append({"run": str(run), "verification": "fail", "verification_errors": ["manifest status is not completed"]})
            else: skipped.append(str(run))
            continue
        reports.append(summarize_run(run))
    result = {"schema": "litchi-0691-trace-summary-v1", "pointer_scope": "per probe process",
              "runs": reports, "skipped": skipped,
              "status": "pass" if reports and all(x["verification"] == "pass" for x in reports) else "fail"}
    output = Path(args.output); output = output if output.is_absolute() else P / output
    dump(output, result)
    print(output)
    return 0 if result["status"] == "pass" else 1

if __name__ == "__main__":
    try: raise SystemExit(main())
    except (OSError, ValueError, KeyError) as exc:
        print(f"trace summary failed: {exc}", file=sys.stderr)
        raise SystemExit(2)
