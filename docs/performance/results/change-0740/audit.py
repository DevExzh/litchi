#!/usr/bin/env python3
"""Independent, read-only audit of the frozen change-0740 FP receipts.

This file deliberately does not import formal.py, contract.py, or parse-stacks.py.
It validates the sealed receipt/build/report identities, decodes archived perf data
when the original file is absent, and recomputes callchain periods from stacks.
"""

from collections import Counter, defaultdict
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys


HERE = Path(__file__).resolve().parent
FORMAL = HERE / "formal"
ROOT = "litchi_perf_baseline::run_pptx_cross_copy_lifecycle"
BASE = "litchi_pptx::opened::cross_copy_plan::"
PLAN = BASE + "plan_cross_slide_copy_for_slides"
PREP = BASE + "prepare_cross_slide_copy_for_slides"
APPLY = (
    BASE + "apply_plan",
    "litchi_pptx::package::model::Package::apply_cross_slide_copy_plan",
)
PHASES = ("plan_ns", "commit_ns", "publication_ns", "reopen_ns", "lifecycle_ns")
BUCKETS = (
    "ambiguous",
    "rooted_apply_excluded",
    "rooted_other",
    "rooted_plan_prepare",
    "unrooted_setup_oracle",
)
EXPECTED = (
    ("pptx_cross_copy_plain_lifecycle", 100),
    ("pptx_cross_copy_media_rich_lifecycle", 10),
    ("pptx_cross_copy_media_rich_lifecycle", 10),
    ("pptx_cross_copy_plain_lifecycle", 100),
    ("pptx_cross_copy_plain_lifecycle", 100),
    ("pptx_cross_copy_media_rich_lifecycle", 10),
)
HEADER = re.compile(
    r"^(?P<comm>\S+)\s+(?P<pid>\d+)/(?P<tid>\d+)\s+"
    r"(?P<time>[0-9]+\.[0-9]+):\s+(?P<period>[0-9]+)\s+cycles:u:\s*$"
)
FRAME = re.compile(r"^\s*(?P<address>[0-9a-fA-F]+)\s+(?P<rest>.+)$")
KNOWN_COMMS = {"taskset", "litchi-perf-bas", "git", "rustc", "uname"}


class AuditError(RuntimeError):
    pass


def need(condition, message):
    if not condition:
        raise AuditError(message)


def load(path):
    with path.open() as stream:
        return json.load(stream)


def sha_bytes(data):
    return hashlib.sha256(data).hexdigest()


def sha_path(path):
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while True:
            chunk = stream.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def decode_gzip(path):
    with gzip.open(path, "rb") as stream:
        return stream.read()


def validate_archive_entry(entry):
    archive = HERE / entry["archive"]
    need(archive.is_file(), f"missing archive {archive}")
    archive_sha, archive_bytes = sha_path(archive)
    need(archive_sha == entry["archive_sha256"], f"archive hash {entry['archive']}")
    need(archive_bytes == entry["archive_bytes"], f"archive size {entry['archive']}")
    decoded = decode_gzip(archive)
    need(sha_bytes(decoded) == entry["raw_sha256"], f"decoded raw hash {entry['raw']}")
    need(len(decoded) == entry["raw_bytes"], f"decoded raw size {entry['raw']}")
    return {
        "archive": entry["archive"],
        "archive_sha256": archive_sha,
        "archive_bytes": archive_bytes,
        "decoded_sha256": sha_bytes(decoded),
        "decoded_bytes": len(decoded),
    }


def validate_receipt_file(name, expected_sha, archives):
    direct = FORMAL / name
    if direct.is_file():
        digest, size = sha_path(direct)
        need(digest == expected_sha, f"receipt hash {name}")
        return {"path": str(direct.relative_to(HERE)), "sha256": digest, "bytes": size, "source": "direct"}
    if name.endswith(".perf.data"):
        raw = "formal/" + name
        entry = archives.get(raw)
        need(entry is not None, f"missing raw/archive receipt {name}")
        info = validate_archive_entry(entry)
        need(info["decoded_sha256"] == expected_sha, f"manifest raw hash {name}")
        return {
            "path": entry["archive"],
            "sha256": expected_sha,
            "bytes": info["decoded_bytes"],
            "source": "gzip-decoded-raw",
            "archive_sha256": info["archive_sha256"],
            "archive_bytes": info["archive_bytes"],
        }
    raise AuditError(f"missing receipt file {name}")


def parse_stack_text(text):
    """Parse perf script records without relying on the frozen parser."""
    need("LOST" not in text and "PERF_RECORD_LOST" not in text, "lost perf record marker")
    records = []
    blocks = [b for b in re.split(r"\n\s*\n", text.strip()) if b.strip()]
    need(blocks, "empty stack stream")
    for block_index, block in enumerate(blocks):
        lines = block.splitlines()
        match = HEADER.fullmatch(lines[0])
        need(match is not None, f"malformed header {block_index}")
        period = int(match.group("period"))
        need(period > 0, f"non-positive period {block_index}")
        frames = []
        for line in lines[1:]:
            match = FRAME.fullmatch(line)
            need(match is not None, f"malformed frame {block_index}")
            rest = match.group("rest")
            split = rest.rfind(" (")
            need(split > 0 and rest.endswith(")"), f"missing dso {block_index}")
            symbol = rest[:split]
            dso = rest[split + 2 : -1]
            need(symbol and dso, f"empty frame field {block_index}")
            frames.append((match.group("address"), symbol, dso))
        need(frames, f"empty frame chain {block_index}")
        records.append({"comm": match_header_comm(lines[0]), "period": period, "frames": frames})
    return records


def match_header_comm(line):
    return HEADER.fullmatch(line).group("comm")


def symbols(record):
    return [frame[1] for frame in record["frames"]]


def classify(record):
    names = symbols(record)
    unknown = any("[unknown]" in frame[1] or "[unknown]" in frame[2] for frame in record["frames"])
    depth_limit = len(names) >= 127
    root = ROOT in names
    plan = PLAN in names
    prep = PREP in names
    apply = any(marker in names for marker in APPLY)
    if unknown or depth_limit or (plan and apply):
        return "ambiguous"
    if not root:
        return "unrooted_setup_oracle"
    if apply:
        return "rooted_apply_excluded"
    if plan and prep:
        if not (names.index(PREP) < names.index(PLAN) < names.index(ROOT)):
            return "ambiguous"
        return "rooted_plan_prepare"
    if prep:
        return "ambiguous"
    return "rooted_other"


def stage(record):
    names = set(symbols(record))
    candidate = BASE + "build_candidate" in names
    writer = "litchi_opc::pkgwriter::PackageWriter::write_to_stream" in names
    generated = "soapberry_zip::preserve::generated_entry" in names
    deflate = "zlib_rs::deflate::deflate" in names
    if candidate and writer and generated and deflate:
        return "candidate_writer_generated_deflate"
    if candidate and writer:
        return "candidate_writer_other"
    if candidate:
        return "candidate_other"
    return "plan_other"


def validate_context(records, binary):
    binary_name = Path(binary).name[:15]
    allowed = KNOWN_COMMS | {binary_name}
    for record_index, record in enumerate(records):
        need(record["comm"] in allowed, f"unexpected comm {record['comm']} at {record_index}")
        for address, name, dso in record["frames"]:
            del address
            if name in (ROOT, PLAN, PREP) or name in APPLY:
                need(dso == binary, f"anchor dso mismatch at {record_index}: {name}")


def analyze_records(records, binary):
    validate_context(records, binary)
    buckets = {name: {"samples": 0, "period": 0} for name in BUCKETS}
    stages = defaultdict(lambda: {"samples": 0, "period": 0})
    leaves = Counter()
    inclusive = Counter()
    total_period = 0
    unknown_count = 0
    depth_count = 0
    root_count = root_period = 0
    plan_count = plan_period = 0
    apply_count = apply_period = 0
    for record in records:
        period = record["period"]
        names = symbols(record)
        bucket = classify(record)
        buckets[bucket]["samples"] += 1
        buckets[bucket]["period"] += period
        total_period += period
        has_unknown = any("[unknown]" in name or "[unknown]" in frame[2] for frame, name in zip(record["frames"], names))
        unknown_count += int(has_unknown)
        depth_count += int(len(names) >= 127)
        root_count += int(ROOT in names)
        root_period += period * int(ROOT in names)
        plan_count += int(PLAN in names)
        plan_period += period * int(PLAN in names)
        apply_count += int(any(marker in names for marker in APPLY))
        apply_period += period * int(any(marker in names for marker in APPLY))
        if bucket == "rooted_plan_prepare":
            current = stage(record)
            stages[current]["samples"] += 1
            stages[current]["period"] += period
            leaves[names[0]] += period
            for name in set(names):
                inclusive[name] += period
    need(sum(item["period"] for item in buckets.values()) == total_period, "bucket period denominator")
    need(sum(item["samples"] for item in buckets.values()) == len(records), "bucket sample denominator")
    strict = stages
    need(sum(item["period"] for item in strict.values()) == buckets["rooted_plan_prepare"]["period"], "stage period denominator")
    need(sum(item["samples"] for item in strict.values()) == buckets["rooted_plan_prepare"]["samples"], "stage sample denominator")
    return {
        "total_samples": len(records),
        "total_period": total_period,
        "unknown_chain_samples": unknown_count,
        "depth_limit_samples": depth_count,
        "root_samples": root_count,
        "root_period": root_period,
        "plan_samples": plan_count,
        "plan_period": plan_period,
        "apply_samples": apply_count,
        "apply_period": apply_period,
        "buckets": buckets,
        "strict_plan_partition": dict(sorted(strict.items())),
        "strict_plan_leaf_period": dict(sorted(leaves.items())),
        "strict_plan_inclusive_symbol_period_overlapping": dict(sorted(inclusive.items())),
    }


def report_contract(report, case, samples, build, source, oracle):
    need(report["schema_version"] == 1, f"report schema {case}")
    need(len(report["results"]) == 1, f"report result count {case}")
    result = report["results"][0]
    need(result["case"] == case == result["corpus"]["name"], f"report case {case}")
    config = report["configuration"]
    need(config["samples_per_case"] == samples, f"report samples {case}")
    need(config["warmup_iterations_per_case"] == 0, f"report warmups {case}")
    need(config["cases"] == [case], f"report case list {case}")
    identity = report["binary_identity"]
    for key in ("path", "binary_sha256", "binary_bytes"):
        need(identity[key] == build["binary"] if key == "path" else identity[key] == build["binary_" + key.split("_", 1)[1]] if key == "binary_sha256" else identity[key] == build["binary_bytes"], f"build identity {key}")
    need(identity["profile"] == "release" and identity["executable"] is True, f"build profile {case}")
    environment = report["environment"]
    need(environment["cpu_affinity"] == "12", f"cpu affinity {case}")
    need(environment["git_revision"] == source["head"], f"git revision {case}")
    need(environment["rustflags"] == build["rustflags"], f"rustflags {case}")
    need(report["tool"]["profile"] == "release", f"tool profile {case}")
    need(report["tool"]["instrumentation"] == "none", f"tool instrumentation {case}")
    expected = oracle[case]
    need(result["corpus"] == expected["corpus"], f"oracle corpus {case}")
    need(result["sink"] == expected["sink"], f"oracle sink {case}")
    cross = result["source"]["pptx_cross_copy"]
    for key, value in expected["cross_copy"].items():
        need(cross.get(key) == value, f"oracle cross_copy {case}:{key}")
    need(result["output_sha256"] == expected["output_sha256"], f"oracle output {case}")
    need(cross["gates"] and all(value is True for value in cross["gates"].values()), f"oracle gates {case}")
    elapsed = result["elapsed_ns"]
    need(sorted(elapsed["sample_order"]) == list(range(samples)), f"sample order {case}")
    need(elapsed["samples"] == sorted(elapsed["samples"]) == cross["lifecycle_ns"], f"lifecycle values {case}")
    for phase in PHASES:
        values = cross[phase]
        need(len(values) == samples and all(type(value) is int and value > 0 for value in values), f"phase {case}:{phase}")
    for index, total in enumerate(cross["lifecycle_ns"]):
        need(sum(cross[name][index] for name in ("plan_ns", "commit_ns", "publication_ns")) <= total, f"phase sum {case}")
    need(result["operation_metrics"]["allocation"]["status"] == "unavailable", f"native allocation {case}")


def expected_commands(index, case, samples, binary):
    stem = f"{index:02d}"
    report = str(FORMAL / f"{stem}.json")
    data = str(FORMAL / f"{stem}.perf.data")
    child = ["taskset", "-c", "12", binary, "--warmup", "0", "--samples", str(samples), "--case", case, "--json", report]
    return (
        ["perf", "record", "-e", "cycles:u", "-F", "997", "--call-graph", "fp,127", "-o", data, "--"] + child,
        ["perf", "script", "--show-lost-events", "-i", data, "-F", "comm,pid,tid,time,event,period,ip,sym,dso"],
    )


def negative_controls(records, binary):
    clean = next(record for record in records if classify(record) == "rooted_plan_prepare")
    base = [dict(frame) if isinstance(frame, dict) else frame for frame in clean["frames"]]

    def altered(change):
        frames = list(base)
        change(frames)
        return {"comm": clean["comm"], "period": clean["period"], "frames": frames}

    controls = []
    controls.append(("unknown_frame", altered(lambda f: f.insert(0, ("1", "[unknown]", "[unknown]"))), "ambiguous"))
    controls.append(("apply_marker_conflict", altered(lambda f: f.insert(0, ("2", APPLY[0], binary))), "ambiguous"))
    controls.append(("root_marker_removed", altered(lambda f: f.__setitem__(next(i for i, x in enumerate(f) if x[1] == ROOT), ("3", "wrong_root", binary))), "unrooted_setup_oracle"))
    controls.append(("order_conflict", altered(lambda f: f.__setitem__(next(i for i, x in enumerate(f) if x[1] == PREP), ("4", PREP, binary))), "rooted_plan_prepare"))
    # The order control is made meaningful by placing prepare after plan.
    order_frames = list(base)
    prep_index = next(i for i, x in enumerate(order_frames) if x[1] == PREP)
    plan_index = next(i for i, x in enumerate(order_frames) if x[1] == PLAN)
    order_frames[prep_index], order_frames[plan_index] = order_frames[plan_index], order_frames[prep_index]
    controls[-1] = ("order_conflict", {"comm": clean["comm"], "period": clean["period"], "frames": order_frames}, "ambiguous")
    results = []
    for name, record, expected in controls:
        actual = classify(record)
        need(actual == expected, f"negative classifier {name}: {actual}")
        results.append({"name": name, "expected": expected, "observed": actual, "passed": True})

    text = render_record(clean)
    malformed = [
        ("bad_header_event", text.replace("cycles:u:", "instructions:u:", 1)),
        ("zero_period", re.sub(r":\s+[0-9]+ cycles:u:", ":          0 cycles:u:", text, count=1)),
        ("malformed_frame", re.sub(r"(?m)^(\s*)[0-9a-fA-F]+ ", r"\1nothex ", text, count=1)),
    ]
    for name, candidate in malformed:
        try:
            parse_stack_text(candidate)
        except AuditError:
            results.append({"name": name, "expected": "reject", "observed": "rejected", "passed": True})
        else:
            raise AuditError(f"negative parser accepted {name}")
    context = dict(clean)
    context["comm"] = "unexpected-process"
    try:
        validate_context([context], binary)
    except AuditError:
        results.append({"name": "unexpected_comm", "expected": "reject", "observed": "rejected", "passed": True})
    else:
        raise AuditError("negative context accepted unexpected_comm")
    dso_frames = list(clean["frames"])
    anchor = next(i for i, item in enumerate(dso_frames) if item[1] == PREP)
    dso_frames[anchor] = (dso_frames[anchor][0], PREP, "/tmp/other-binary")
    context = {"comm": clean["comm"], "period": clean["period"], "frames": dso_frames}
    try:
        validate_context([context], binary)
    except AuditError:
        results.append({"name": "anchor_wrong_dso", "expected": "reject", "observed": "rejected", "passed": True})
    else:
        raise AuditError("negative context accepted anchor_wrong_dso")
    return results


def render_record(record):
    lines = [f"{record['comm']} 1/1 1.0: {record['period']} cycles:u: "]
    lines.extend(f"\t{address} {symbol} ({dso})" for address, symbol, dso in record["frames"])
    return "\n".join(lines)


def main():
    build = load(HERE / "build-fp.json")
    source = load(HERE / "source.json")
    oracle = load(HERE / "oracle.json")
    freeze = load(HERE / "freeze.json")
    need(build["exit"] == 0, "FP build did not pass")
    need(sha_path(HERE / "build-fp.json")[0] == freeze["build_sha256"], "frozen build receipt hash")
    frozen_hashes = {}
    for name, expected_hash in freeze["scripts"].items():
        actual_hash, _ = sha_path(HERE / name)
        need(actual_hash == expected_hash, f"frozen script hash {name}")
        frozen_hashes[name] = actual_hash
    binary = Path(build["binary"])
    if binary.is_file():
        binary_hash, binary_bytes = sha_path(binary)
    else:
        cleanup = load(HERE / "cleanup.json")
        witnesses = [row for row in cleanup["removed"] if row["binary"] == str(binary)]
        need(len(witnesses) == 1, "missing exact FP cleanup witness")
        witness = witnesses[0]
        need(Path(witness["removed"]) == binary.parent.parent and not Path(witness["removed"]).exists(), "FP cleanup path")
        need(not Path(witness["perf_cache"]["removed"]).exists(), "FP perf cache remains")
        binary_hash, binary_bytes = witness["sha256"], witness["bytes"]
    need(binary_hash == build["binary_sha256"] and binary_bytes == build["binary_bytes"], "FP binary identity")

    archive_entries = load(HERE / "raw-archives.json")
    archives = {entry["raw"]: entry for entry in archive_entries}
    formal_raw = {f"formal/{index:02d}.perf.data" for index in range(6)}
    need(formal_raw <= archives.keys(), "formal archive map incomplete")
    archive_checks = {raw: validate_archive_entry(archives[raw]) for raw in sorted(formal_raw)}

    manifest = load(FORMAL / "manifest.json")
    need(len(manifest) == len(EXPECTED), "formal receipt count")
    rows = []
    all_records = []
    analysis = load(HERE / "analysis.json")
    analysis_by_index = {row["index"]: row for row in analysis["rows"]}
    for index, (case, samples) in enumerate(EXPECTED):
        receipt = manifest[index]
        need(receipt["mode"] == "profile" and receipt["case"] == case, f"receipt schedule {index}")
        need(receipt["samples"] == samples and receipt["warmups"] == 0, f"receipt parameters {index}")
        record_command, script_command = expected_commands(index, case, samples, build["binary"])
        need(receipt["command"] == record_command, f"record command {index}")
        need(receipt["script_command"] == script_command, f"script command {index}")
        need(receipt["exit"] == 0 and receipt["script_exit"] == 0, f"receipt exit {index}")
        need(receipt["ended"] >= receipt["started"], f"receipt times {index}")
        files = receipt["files"]
        expected_files = {f"{index:02d}.{suffix}" for suffix in ("stdout", "stderr", "script.stderr", "stacks", "perf.data", "json")}
        need(set(files) == expected_files, f"receipt file list {index}")
        receipt_checks = {name: validate_receipt_file(name, digest, archives) for name, digest in sorted(files.items())}
        need((FORMAL / f"{index:02d}.stdout").read_bytes() == b"", f"stdout {index}")
        need((FORMAL / f"{index:02d}.script.stderr").read_bytes() == b"", f"perf script stderr {index}")
        stderr = (FORMAL / f"{index:02d}.stderr").read_text().lower()
        need(not any(word in stderr for word in ("failed", "error:", "corruption", "lost")), f"perf stderr {index}")
        report = load(FORMAL / f"{index:02d}.json")
        report_contract(report, case, samples, build, source, oracle)
        stack_path = FORMAL / f"{index:02d}.stacks"
        stack_bytes = stack_path.read_bytes()
        stack_sha = sha_bytes(stack_bytes)
        need(stack_sha == files[f"{index:02d}.stacks"], f"stack receipt hash {index}")
        records = parse_stack_text(stack_bytes.decode())
        stats = analyze_records(records, build["binary"])
        expected_analysis = analysis_by_index[index]
        for key in ("total_samples", "total_period", "unknown_chain_samples", "depth_limit_samples", "buckets", "strict_plan_partition", "strict_plan_leaf_period", "strict_plan_inclusive_symbol_period_overlapping"):
            need(stats[key] == expected_analysis[key], f"independent analysis mismatch {index}:{key}")
        all_records.extend(records)
        rows.append({
            "index": index,
            "case": case,
            "samples": samples,
            "report_sha256": files[f"{index:02d}.json"],
            "stack_sha256": stack_sha,
            "stack_bytes": len(stack_bytes),
            "receipt_files": receipt_checks,
            "analysis": stats,
            "analysis_matches_frozen_output": True,
        })

    aggregate = analyze_records(all_records, build["binary"])
    negatives = negative_controls(all_records, build["binary"])
    output = {
        "status": "passed",
        "scope": "six formal FP profiles; period attribution only; no causal or removable fraction",
        "independence": {
            "parser": "self-contained audit parser; frozen parse-stacks.py was not imported",
            "frozen_script_sha256": frozen_hashes,
            "build_receipt_sha256": freeze["build_sha256"],
            "binary_sha256": binary_hash,
            "binary_bytes": binary_bytes,
        },
        "schedule": [{"index": i, "case": case, "samples": samples} for i, (case, samples) in enumerate(EXPECTED)],
        "archive_replay": {raw: archive_checks[raw] for raw in sorted(archive_checks)},
        "reports": rows,
        "aggregate": aggregate,
        "negative_controls": negatives,
        "checks": {
            "formal_receipts": True,
            "report_oracles_and_build_ids": True,
            "raw_receipt_hashes": True,
            "gzip_decoded_raw_hashes": True,
            "period_denominators": True,
            "bucket_and_stage_recomputation": True,
            "context_and_malformed_rejection": True,
        },
    }
    (HERE / "audit.json").write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")
    print("PASS audit: six formal receipts, independent stacks, oracle/build/hash checks")
    print(f"PASS audit: samples={aggregate['total_samples']} period={aggregate['total_period']} strict_plan_period={aggregate['buckets']['rooted_plan_prepare']['period']}")
    print(f"PASS audit: negative_controls={len(negatives)} gzip_replays={len(archive_checks)}")


if __name__ == "__main__":
    try:
        main()
    except (AuditError, AssertionError, KeyError, OSError, ValueError) as error:
        print(f"FAIL audit: {error}", file=sys.stderr)
        raise
