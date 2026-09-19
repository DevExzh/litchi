#!/usr/bin/env python3
"""Verify and summarize completed 0693 temporary trace runs."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
RUNS = P / "trace-runs"
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
    if not words or not words[0].startswith("LITCHI0693_"):
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
    restoration_path = run / "restoration.json"
    restoration = manifest.get("restoration", [])
    if not restoration or not restoration_path.is_file():
        errors.append("manifest/restoration archive is missing")
    elif load(restoration_path) != restoration:
        errors.append("restoration.json differs from manifest")
    for item in restoration:
        if item.get("restored_exact") is not True:
            errors.append(f"restoration failed for {item.get('path')}")
        path = ROOT / item["path"]
        if not path.is_file() or sha(path) != item.get("restored_sha256"):
            errors.append(f"live source hash differs after restore: {item.get('path')}")
    hashes_path = run / "source-hashes.json"
    patch_path = run / "trace.patch"
    if not hashes_path.is_file() or not patch_path.is_file():
        errors.append("source hash archive or trace.patch is missing")
    else:
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
            else:
                if restored.get("current_sha256") != item.get("after_sha256"):
                    errors.append(f"patched hash does not equal archived after for {item['path']}")
                if restored.get("restored_sha256") != item.get("before_sha256"):
                    errors.append(f"restore hash does not equal archived before for {item['path']}")
    records_path = run / "probe-runs.json"
    records = load(records_path) if records_path.is_file() else []
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
    if not binary or not manifest.get("binary_sha256"):
        errors.append("manifest has no trace binary binding")
    else:
        path = Path(binary)
        if not path.is_absolute():
            path = ROOT / path
        if path.is_file():
            if sha(path) != manifest["binary_sha256"]:
                errors.append("trace binary hash differs from manifest")
        else:
            cleanup_path = P / "cleanup.json"
            removed = load(cleanup_path).get("removed", []) if cleanup_path.is_file() else []
            owned_target = ROOT.parent / "litchi-target-0693"
            authorized_removal = any(
                item.get("removed") is True
                and item.get("kind") == "directory"
                and Path(item["path"]) == owned_target
                for item in removed
            )
            if not authorized_removal or not path.is_relative_to(owned_target):
                errors.append("trace binary missing without owned-target cleanup receipt")
    return errors


def event_interval(operation: str, ordinal: int, total: int) -> str:
    if operation == "apply":
        if ordinal == 0:
            return "derive_target_setup"
        if ordinal == total - 2:
            return "route_source"
        if ordinal == total - 1:
            return "route_commit"
        return "unexpected_intermediate"
    return "measured_capture"


def summarize_events(codec: list[dict], parts: list[dict]) -> dict:
    uri_by_key = defaultdict(set)
    stages_by_key = defaultdict(set)
    for part in parts:
        key = (part["raw_ptr"], part["raw_len"])
        uri_by_key[key].add(part["uri"])
        stages_by_key[key].add(part["stage"])
    pointers = {}
    modes = Counter()
    statuses = Counter()
    inputs = outputs = 0
    for call in codec:
        key = (call["input_ptr"], call["input_len"])
        row_key = f"0x{key[0]:x}/{key[1]}"
        row = pointers.setdefault(row_key, {
            "input_ptr": f"0x{key[0]:x}", "input_len": key[1], "calls": 0,
            "input_bytes": 0, "output_bytes": 0, "mode_counts": {},
            "status_counts": {}, "output_ptrs": [], "mapped_uris": [],
            "part_stages": [],
        })
        row["calls"] += 1
        row["input_bytes"] += call["input_len"]
        row["output_bytes"] += call["output_len"]
        row["mode_counts"][call["mode"]] = row["mode_counts"].get(call["mode"], 0) + 1
        row["status_counts"][call["status"]] = row["status_counts"].get(call["status"], 0) + 1
        output_ptr = f"0x{call['output_ptr']:x}"
        if output_ptr not in row["output_ptrs"]:
            row["output_ptrs"].append(output_ptr)
        row["mapped_uris"] = sorted(uri_by_key.get(key, ()))
        row["part_stages"] = sorted(stages_by_key.get(key, ()))
        modes[call["mode"]] += 1
        statuses[call["status"]] += 1
        inputs += call["input_len"]
        outputs += call["output_len"]
    unmapped = [row for row in pointers.values() if not row["mapped_uris"]]
    return {
        "codec_calls": len(codec), "input_bytes": inputs, "output_bytes": outputs,
        "borrowed_calls": modes.get("borrow", 0), "owned_calls": modes.get("owned", 0),
        "error_calls": statuses.get("error", 0), "mode_counts": dict(modes),
        "status_counts": dict(statuses), "part_records": len(parts),
        "processed_xml_part_records": sum(x.get("stage") == "processed_xml" for x in parts),
        "source_bound_part_records": sum(x.get("stage") == "source_bound" for x in parts),
        "mapped_part_records": sum(bool(uri_by_key.get((x["raw_ptr"], x["raw_len"]))) for x in parts),
        "unmapped_codec_calls": sum(x["calls"] for x in unmapped),
        "raw_pointer_len": sorted(pointers.values(), key=lambda x: (x["input_ptr"], x["input_len"])),
        "pointer_scope": "one probe process; do not compare addresses across stderr files",
    }


def parse_trace(path: Path, operation: str, require_model: bool) -> dict:
    intervals = []
    outside = {"codec": [], "parts": [], "layout": []}
    all_events = {"codec": [], "parts": []}
    active = None
    errors = []
    for line_no, line in enumerate(path.read_text(errors="replace").splitlines(), 1):
        prefix, values = fields(line)
        if prefix == "LITCHI0693_MODEL":
            stage = values.get("stage")
            if stage == "capture_begin":
                if active is not None:
                    errors.append(f"{path.name}:{line_no}: nested capture_begin")
                active = {"begin_line": line_no, "codec": [], "parts": [], "layout": []}
            elif stage == "capture_success":
                if active is None:
                    errors.append(f"{path.name}:{line_no}: capture_success without begin")
                else:
                    active["end_line"] = line_no
                    intervals.append(active)
                    active = None
            elif stage == "capture_layout":
                target = active if active is not None else outside
                try:
                    target["layout"].append({
                        "line": line_no,
                        "proof_entry_bytes": int(values["proof_entry_bytes"]),
                        "slide_entry_bytes": int(values["slide_entry_bytes"]),
                    })
                except (KeyError, ValueError) as exc:
                    errors.append(f"{path.name}:{line_no}: malformed capture layout record: {exc}")
            continue
        target = active if active is not None else outside
        try:
            if prefix == "LITCHI0693_CODEC":
                event = {"line": line_no, "call": int(values["call"]),
                         "input_ptr": integer(values["input_ptr"]), "input_len": int(values["input_len"]),
                         "output_ptr": integer(values["output_ptr"]), "output_len": int(values["output_len"]),
                         "mode": values["mode"], "status": values["status"]}
                target["codec"].append(event)
                all_events["codec"].append(event)
            elif prefix == "LITCHI0693_PART":
                event = {"line": line_no, "stage": values["stage"], "uri": values["uri"],
                         "raw_ptr": integer(values["raw_ptr"]), "raw_len": int(values["raw_len"])}
                target["parts"].append(event)
                all_events["parts"].append(event)
        except (KeyError, ValueError) as exc:
            if prefix in ("LITCHI0693_CODEC", "LITCHI0693_PART"):
                errors.append(f"{path.name}:{line_no}: malformed {prefix} record: {exc}")
    if active is not None:
        errors.append(f"{path.name}: capture_begin has no capture_success")
        intervals.append(active)
    if require_model and not intervals:
        errors.append(f"{path.name}: no capture boundaries")
    summaries = []
    total = len(intervals)
    for ordinal, item in enumerate(intervals):
        if require_model and len(item["layout"]) != 1:
            errors.append(
                f"{path.name}: capture interval {ordinal} has "
                f"{len(item['layout'])} layout records, expected one"
            )
        summary = summarize_events(item["codec"], item["parts"])
        summary.update({"ordinal": ordinal, "role": event_interval(operation, ordinal, total),
                        "begin_line": item["begin_line"], "end_line": item.get("end_line"),
                        "call_first": item["codec"][0]["call"] if item["codec"] else None,
                        "call_last": item["codec"][-1]["call"] if item["codec"] else None,
                        "layout": item["layout"]})
        summaries.append(summary)
    outside_summary = summarize_events(outside["codec"], outside["parts"])
    outside_summary["layout"] = outside["layout"]
    return {"intervals": summaries, "outside_capture": outside_summary,
            "process": summarize_events(all_events["codec"], all_events["parts"]), "errors": errors}


def summarize_run(run: Path) -> dict:
    manifest = load(run / "manifest.json")
    errors = verify_files(run, manifest)
    legs = []
    for record in manifest.get("probe_runs", []):
        label = f"{record['case']}-{record['operation']}-{record['count']}"
        stderr = run / (label + ".stderr")
        if not stderr.is_file():
            errors.append(f"{label}: stderr is missing")
            continue
        parsed = parse_trace(stderr, record["operation"], bool(manifest.get("model_trace")))
        errors.extend(f"{label}: {error}" for error in parsed["errors"])
        legs.append({"case": record["case"], "source": record["source"],
                     "operation": record["operation"], "count": record["count"],
                     "stderr": stderr.name, "intervals": parsed["intervals"],
                     "outside_capture": parsed["outside_capture"], "process": parsed["process"]})
    return {"run": str(run), "profile": manifest.get("profile"),
            "model_trace": manifest.get("model_trace"), "verification_errors": errors,
            "verification": "pass" if not errors else "fail", "legs": legs}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_dirs", nargs="*", help="completed trace run directories")
    parser.add_argument("--output", default=str(P / "trace-summary.json"))
    args = parser.parse_args()
    paths = [Path(x).resolve() for x in args.run_dirs] if args.run_dirs else sorted(
        (x for x in RUNS.iterdir() if x.is_dir()), key=lambda x: x.name) if RUNS.is_dir() else []
    reports, skipped = [], []
    for run in paths:
        manifest_path = run / "manifest.json"
        if not manifest_path.is_file():
            if args.run_dirs:
                reports.append({"run": str(run), "verification": "fail",
                                "verification_errors": ["manifest is missing"]})
            else:
                skipped.append(str(run))
            continue
        manifest = load(manifest_path)
        if manifest.get("status") != "completed":
            if args.run_dirs:
                reports.append({"run": str(run), "verification": "fail",
                                "verification_errors": ["manifest status is not completed"]})
            else:
                skipped.append(str(run))
            continue
        reports.append(summarize_run(run))
    result = {"schema": "litchi-0693-trace-summary-v1", "pointer_scope": "per probe process",
              "runs": reports, "skipped": skipped,
              "status": "pass" if reports and all(x["verification"] == "pass" for x in reports) else "fail"}
    output = Path(args.output)
    output = output if output.is_absolute() else P / output
    dump(output, result)
    print(output)
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, KeyError) as exc:
        print(f"trace summary failed: {exc}", file=sys.stderr)
        raise SystemExit(2)
