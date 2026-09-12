"""Read-only, bounded verifier for the 0530 XLSX planning attribution."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

H = Path(__file__).resolve().parent
R = H.parents[3]
P29 = H.parent / "change-0529"
PLAN, RUN, FROZEN = H / "plan.json", H / "run.py", H / "frozen-inputs.json"
BASE = H / "baseline"
PRIOR_MANIFEST = P29 / "final" / "source-manifest.json"
PRIOR_QUALITY, PRIOR_SEAL = P29 / "quality-summary.json", P29 / "SHA256SUMS"
SCRATCH = Path("/tmp/litchi-goal-0530")
TARGET = Path("/home/zhuhe/litchi-goal-0530-target")
OWNED = [str(SCRATCH), str(TARGET)]
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SHAPES = ["medium", "dense-sparse"]
OWNERS = [
    "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets",
    "litchi_xlsx::cell_values::snapshot::MultiSnapshot::load_source_backed",
]
ANALYZER = H / "analyze_planning.py"
REPORT_NAMES = ("planning-analysis.json", "planning-comparison.json",
                "profile-analysis.json", "analysis.json")


class EvidenceError(ValueError):
    pass


class Pending(EvidenceError):
    pass


def req(ok: bool, message: str) -> None:
    if not ok:
        raise EvidenceError(message)


def need(path: Path, label: str | None = None) -> Path:
    if not path.exists():
        raise Pending(f"{label or path} is missing")
    req(not path.is_symlink(), f"{label or path} is a symlink")
    return path


def read(path: Path) -> Any:
    need(path)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def digest(path: Path) -> str:
    try:
        h = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                h.update(block)
        return h.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def valid_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe(value: Any, label: str) -> str:
    req(isinstance(value, str) and value and not Path(value).is_absolute(),
        f"{label} is not a safe relative path")
    path = Path(value)
    req(path.as_posix() == value and ".." not in path.parts,
        f"{label} escapes its root")
    return value


def times(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    def parse(item: Any, suffix: str) -> dt.datetime:
        req(isinstance(item, str), f"{label}.{suffix} is not a timestamp")
        try:
            parsed = dt.datetime.fromisoformat(item)
        except ValueError as error:
            raise EvidenceError(f"{label}.{suffix} is invalid") from error
        req(parsed.tzinfo is not None, f"{label}.{suffix} has no timezone")
        return parsed
    start = parse(value.get("start_utc"), "start_utc")
    end = parse(value.get("end_utc"), "end_utc")
    seconds = value.get("seconds")
    req(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
        and math.isfinite(float(seconds)) and seconds > 0 and end > start,
        f"{label} interval is invalid")
    return start, end


def plan() -> dict[str, Any]:
    frozen = read(FROZEN)
    req(isinstance(frozen, dict) and set(frozen) == {"plan.json", "run.py"}
        and frozen["plan.json"] == digest(PLAN) and frozen["run.py"] == digest(RUN),
        "frozen plan/driver binding differs")
    value = read(PLAN)
    req(value.get("status") == "frozen-before-build-and-capture"
        and isinstance(value.get("revision"), str)
        and re.fullmatch(r"[0-9a-f]{40}", value["revision"]), "plan is not frozen")
    req(read(H / "start-state.json").get("revision") == value["revision"],
        "start-state revision differs")
    req(value.get("priority") == "OLE2/OOXML active; ODF deferred; iWork excluded",
        "plan priority differs")
    primary, profile = value.get("primary"), value.get("profile")
    req(isinstance(primary, dict) and primary.get("case") == CASE
        and primary.get("shapes") == SHAPES, "primary plan differs")
    req(isinstance(profile, dict) and profile.get("shapes") == SHAPES
        and profile.get("repeats") == 2 and profile.get("warmup") == 0
        and profile.get("samples") == 1 and profile.get("owner_candidates") == OWNERS,
        "profile plan differs")
    req(value.get("candidate_source_roots") == ["crates/litchi-xlsx/"]
        and value.get("capture_lanes") == ["build-normal", "profile"]
        and value.get("owned_paths") == OWNED, "scope or owned paths differ")
    req(not any(key in value for key in ("allocation", "hardware", "guards", "candidate")),
        "0530 unexpectedly contains numerical/candidate lanes")
    req("attribution" in str(value.get("admission", "")).lower()
        and "no runtime" in str(value.get("hypothesis", "")).lower(),
        "attribution-only scope is not frozen")
    return value


def manifest(path: Path) -> dict[str, str]:
    value = read(path)
    req(isinstance(value, dict) and value, f"{path} is not a source manifest")
    result = {}
    for name, value in value.items():
        safe(name, f"{path} source path")
        req((name.startswith("crates/") or name.startswith("tools/perf-baseline/")
             or name.startswith(".cargo/") or name in
             ("Cargo.toml", "Cargo.lock", "rust-toolchain.toml"))
            and valid_digest(value), f"{path} has an invalid source entry: {name}")
        req(name not in result, f"{path} repeats {name}")
        result[name] = value
    return result


def current_sources() -> set[str]:
    tracked = subprocess.check_output([
        "git", "ls-files", "-z", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml"], cwd=R).split(b"\0")
    extra = subprocess.check_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline"], cwd=R).split(b"\0")
    names = {x.decode() for x in tracked if x}
    names.update(x.decode() for x in extra if x.endswith(b".rs"))
    return {name for name in names if (R / name).is_file()}


def validate_source() -> dict[str, Any]:
    plan()
    current = manifest(BASE / "source-manifest.json")
    prior = manifest(PRIOR_MANIFEST)
    req((BASE / "source-manifest.json").read_bytes() == PRIOR_MANIFEST.read_bytes()
        and current == prior, "baseline is not byte-equal to 0529 final")
    req(need(BASE / "source.patch").read_bytes() == b"", "baseline source patch is not empty")
    req(not (H / "candidate").exists(), "candidate stage is out of scope for 0530")
    req(current_sources() == set(current), "current source inventory differs")
    for name, expected in current.items():
        path = R / name
        req(path.is_file() and not path.is_symlink() and digest(path) == expected,
            f"source hash differs: {name}")
    adr = read(H / "adr-manifest.json").get("files")
    req(isinstance(adr, dict) and adr, "ADR manifest is missing")
    for name, expected in adr.items():
        safe(name, "ADR path")
        req(valid_digest(expected) and digest(R / name) == expected, f"ADR hash differs: {name}")
    return {"manifest_sha256": digest(BASE / "source-manifest.json"),
            "manifest_entries": len(current), "prior_final_equal": True, "source_patch": "empty"}


def validate_seal(path: Path, root: Path) -> int:
    expected: dict[str, str] = {}
    for line in need(path).read_text(encoding="utf-8").splitlines():
        fields = line.split("  ", 1)
        req(len(fields) == 2 and valid_digest(fields[0]), f"malformed seal line: {line}")
        name = safe(fields[1], "seal path")
        req(name != path.name and name not in expected, "seal repeats or contains itself")
        target = root / name
        req(target.is_file() and not target.is_symlink(), f"sealed file is missing: {name}")
        expected[name] = fields[0]
    actual = {item.relative_to(root).as_posix(): digest(item) for item in root.rglob("*")
              if item.is_file() and not item.is_symlink() and item != path}
    req(actual == expected and not any(item.is_symlink() for item in root.rglob("*")),
        "sealed inventory is not exact")
    return len(expected)


def validate_prior() -> dict[str, Any]:
    reuse = read(H / "quality-reuse.json")
    source_sha = digest(BASE / "source-manifest.json")
    req(reuse.get("status") == "pass" and reuse.get("source_manifest_sha256") == source_sha
        and reuse.get("prior_quality_summary") == "docs/performance/results/change-0529/quality-summary.json"
        and reuse.get("prior_quality_summary_sha256") == digest(PRIOR_QUALITY)
        and reuse.get("prior_seal_sha256") == digest(PRIOR_SEAL)
        and reuse.get("checks") == 14 and reuse.get("successful_test_executions") == 2024,
        "quality reuse binding differs")
    quality, total = read(PRIOR_QUALITY), 0
    rows = quality.get("checks")
    req(quality.get("status") == "pass" and quality.get("stage") == "final"
        and isinstance(rows, list) and len(rows) == 14, "0529 quality summary differs")
    for row in rows:
        receipt = P29 / "final" / row["name"]
        req(digest(receipt) == row["receipt_sha256"] and row["exit_code"] == 0,
            f"0529 quality receipt differs: {receipt}")
        log = receipt.with_name(receipt.name.replace(".receipt.json", ".stdout"))
        count = sum(int(x) for x in re.findall(r"test result: ok\. (\d+) passed;",
                                                 log.read_text(encoding="utf-8")))
        req(count == row["executed_tests"], f"0529 quality count differs: {receipt}")
        total += count
    req(total == quality.get("executed_tests") == 2024, "0529 quality total differs")
    return {"checks": 14, "successful_test_executions": total,
            "prior_seal_entries": validate_seal(PRIOR_SEAL, P29)}


def common(path: Path, source_sha: str, binary_sha: str | None = None) -> dict[str, Any]:
    value = read(path)
    start, end = times(value, str(path))
    req(value.get("exit_code") == 0 and value.get("plan_sha256") == digest(PLAN)
        and value.get("script_sha256") == digest(RUN)
        and value.get("source_manifest_sha256") == source_sha
        and value.get("working_source_manifest_sha256") == source_sha,
        f"{path} receipt binding differs")
    req(value.get("environment", {}).get("TMPDIR") == str(TARGET / "test-tmp"),
        f"{path} TMPDIR differs")
    if binary_sha is not None:
        req(value.get("binary_sha256") == binary_sha, f"{path} binary differs")
    return {"path": str(path.relative_to(H)), "start": start, "end": end,
            "sha256": digest(path), "value": value}


def artifacts(value: dict[str, Any], folder: Path, name: str, profile: bool = False) -> None:
    items = value.get("artifacts")
    req(isinstance(items, dict) and {name + ".stdout", name + ".stderr"} <= set(items),
        f"{name} artifact inventory is incomplete")
    if profile:
        req(name + ".json" in items and any(x.startswith(name + ".callgrind") for x in items),
            f"{name} profile artifacts are incomplete")
    else:
        req(set(items) == {name + ".stdout", name + ".stderr"}, f"{name} artifacts differ")
    for file, expected in items.items():
        safe(file, f"{name} artifact")
        target = folder / file
        req(Path(file).name == file and target.is_file() and not target.is_symlink()
            and valid_digest(expected) and digest(target) == expected,
            f"{name} artifact hash differs: {file}")


def validate_build() -> tuple[dict[str, Any], dict[str, Any]]:
    p = plan()
    source_sha = digest(BASE / "source-manifest.json")
    receipt = BASE / "build-normal.receipt.json"
    row = common(receipt, source_sha)
    command = ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
               "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
               "--bin", "litchi-perf-baseline", "--target-dir", str(TARGET)]
    req(row["value"].get("binary_sha256") is None and row["value"].get("command") == command,
        "build-normal command differs")
    artifacts(row["value"], BASE, "build-normal")
    identity = read(BASE / "binary-normal.json")
    binary = SCRATCH / "baseline-normal"
    req(identity.get("path") == str(binary) and valid_digest(identity.get("sha256"))
        and isinstance(identity.get("bytes"), int) and identity["bytes"] > 0
        and identity.get("source_manifest_sha256") == source_sha
        and identity.get("build_receipt_sha256") == row["sha256"], "binary identity differs")
    if binary.exists():
        req(binary.is_file() and not binary.is_symlink() and digest(binary) == identity["sha256"]
            and binary.stat().st_size == identity["bytes"], "binary custody differs")
    else:
        cleanup = read(H / "cleanup.json")
        req(cleanup.get("owned_paths_absent") is True and cleanup.get("removed") == OWNED,
            "missing binary is not cleanup-bound")
    storage = read(H / "storage.json")
    req(storage.get("scratch_symlink") == str(SCRATCH)
        and storage.get("target") == str(TARGET / "retained-binaries"), "storage differs")
    return row, {"path": str(binary), "sha256": identity["sha256"], "bytes": identity["bytes"]}


def symbol(binary: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    source_sha = digest(BASE / "source-manifest.json")
    receipt_path = BASE / "symbols.receipt.json"
    receipt = common(receipt_path, source_sha, binary["sha256"])
    value = receipt["value"]
    req(value.get("command") == ["nm", "-C", binary["path"]], "nm command differs")
    artifacts(value, BASE, "symbols")
    observation = read(H / "symbol-observation.json")
    candidates = observation.get("candidates")
    req(observation.get("plan_sha256") == digest(PLAN)
        and observation.get("binary_sha256") == binary["sha256"]
        and isinstance(candidates, dict) and list(candidates) == OWNERS,
        "symbol observation binding differs")
    owner = observation.get("owner")
    req(owner in OWNERS, "symbol owner is not planned")
    output = BASE / "symbols.stdout"
    req(observation.get("stdout_sha256") == digest(output), "nm stdout hash differs")
    raw = output.read_text(encoding="utf-8")
    def present(candidate: str) -> bool:
        return re.search(r"(?<![A-Za-z0-9_:])" + re.escape(candidate) + r"(?:\(|\s|$)",
                         raw, re.MULTILINE) is not None
    first = next((candidate for candidate in OWNERS if present(candidate)), None)
    req(first == owner, "selected owner is not the first exact nm -C match")
    req(all(isinstance(lines, list) and lines for lines in candidates.values()),
        "symbol candidate observations are missing")
    req(any(line in raw for line in candidates[owner]), "selected nm candidate is absent")
    observation_receipt = observation.get("receipt_sha256")
    req(observation_receipt == receipt["sha256"], "symbol receipt binding differs")
    return {"owner": owner, "first_match": first, "nm_sha256": digest(output)}, receipt


def raw_helper() -> Any:
    path = P29 / "analyze.py"
    spec = importlib.util.spec_from_file_location("litchi_0530_raw_validator", path)
    req(spec is not None and spec.loader is not None, "0529 raw validator cannot load")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.BASE


def validated_profile(path: Path, name: str, shape: str,
                      binary: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    helper = raw_helper()
    job = {"name": name, "kind": "primary", "guard": None, "repeat": int(name.split("-")[1][1:]),
           "case": CASE, "shape": shape, "warmup": 0, "samples": 1}
    return helper.validate_result(read(path), p, job, binary, False)


def normalized_identity(row: dict[str, Any]) -> dict[str, Any]:
    identity = json.loads(json.dumps(row["identity"]))
    config = identity["configuration"]
    config["samples_per_case"], config["warmup_iterations_per_case"] = "planned", "planned"
    return identity


def reference_identity(shape: str) -> dict[str, Any]:
    helper = raw_helper()
    p = read(P29 / "plan.json")
    binary = read(P29 / "baseline" / "binary-normal.json")
    job = {"name": "reference", "kind": "primary", "guard": None, "repeat": 1,
           "case": CASE, "shape": shape, "warmup": 20, "samples": 200}
    raw = read(P29 / "baseline" / f"native-r1-primary-{shape}.json")
    return normalized_identity(helper.validate_result(raw, p, job, binary, False))


def validate_profiles() -> dict[str, Any]:
    p = plan()
    build_row, binary = validate_build()
    selected, symbol_row = symbol(binary)
    source_sha = digest(BASE / "source-manifest.json")
    expected = {f"profile-r{repeat}-{shape}" for repeat in (1, 2) for shape in SHAPES}
    actual = {x.name.removesuffix(".receipt.json") for x in BASE.glob("profile-*.receipt.json")}
    req(actual == expected, f"profile receipt inventory differs: {sorted(actual)}")
    receipt_names = {x.name.removesuffix(".receipt.json")
                     for x in BASE.glob("*.receipt.json")}
    req(receipt_names == {"build-normal", "symbols"} | expected,
        "0530 contains an unplanned native/allocator or capture receipt")
    rows = []
    references = {shape: reference_identity(shape) for shape in SHAPES}
    for name in sorted(expected):
        shape = "-".join(name.split("-")[2:])
        path = BASE / (name + ".receipt.json")
        row = common(path, source_sha, binary["sha256"])
        command = ["taskset", "-c", str(p["cpu"]), "valgrind", "--tool=callgrind",
                   "--collect-atstart=no", "--toggle-collect=" + selected["owner"],
                   "--zero-before=" + selected["owner"], "--dump-after=" + selected["owner"],
                   "--callgrind-out-file=" + str(BASE / (name + ".callgrind")), binary["path"],
                   "--warmup", "0", "--samples", "1", "--case", CASE,
                   "--xlsx-cell-crud-shape", shape, "--json", str(BASE / (name + ".json"))]
        req(row["value"].get("command") == command, f"{path} command differs")
        artifacts(row["value"], BASE, name, profile=True)
        validated = validated_profile(BASE / (name + ".json"), name, shape, binary, p)
        req(normalized_identity(validated) == references[shape],
            f"{path} corpus/sink/source identity differs from 0529 native reference")
        row["name"] = name
        rows.append(row)
    ordered = sorted([build_row, symbol_row] + rows, key=lambda row: row["start"])
    for left, right in zip(ordered, ordered[1:]):
        req(left["end"] <= right["start"], f"receipt intervals overlap: {left['path']} and {right['path']}")
    return {"profile_receipts": len(rows), "serial_receipts": len(ordered),
            "binary_sha256": binary["sha256"], "symbol": selected}


def validate_context() -> dict[str, Any]:
    import importlib.util
    path = H / "analyze_context.py"
    spec = importlib.util.spec_from_file_location("planning_context_replay", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    value = module.analyze()
    req(read(H / "planning-context.json") == value, "planning context replay differs")
    req((H / "planning-context.md").read_text() == module.markdown(value),
        "planning context Markdown replay differs")
    return {"json_sha256": digest(H / "planning-context.json"),
            "script_sha256": digest(path), "exact_replay": True}


def validate_draft() -> dict[str, Any]:
    value = read(H / "next-pilot-source.json")
    name = "crates/litchi-ooxml-common/src/mce/codec.rs"
    req(value.get("status") == "unapplied-unmeasured-draft" and value.get("file") == name,
        "future draft scope differs")
    source = (R / name).read_text()
    req(digest(R / name) == value.get("baseline_sha256"), "future draft base differs")
    old = "    if !xml\n        .windows(NAMESPACE.len())\n        .any(|w| w == NAMESPACE.as_bytes())\n    {"
    new = "    if memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_none() {"
    req(source.count(old) == 1, "namespace predicate base differs")
    candidate = source.replace(old, new)
    req(hashlib.sha256(candidate.encode()).hexdigest() == value.get("draft_sha256"),
        "future draft source differs")
    import difflib
    patch = "".join(difflib.unified_diff(source.splitlines(True), candidate.splitlines(True),
                      fromfile="a/" + name, tofile="b/" + name))
    path = H / "namespace-search.patch"
    req(path.read_text() == patch and digest(path) == value.get("patch_sha256"),
        "future draft patch differs")
    return {"status": "unapplied-unmeasured-draft", "patch_sha256": digest(path),
            "exact_predicate_substitution": True}


def report_path() -> Path:
    paths = [H / name for name in REPORT_NAMES if (H / name).is_file()]
    req(len(paths) <= 1, "multiple planning reports are present")
    return need(paths[0] if paths else H / REPORT_NAMES[0], "planning analyzer report")


def validate_analysis() -> dict[str, Any]:
    validate_profiles()
    need(ANALYZER, "analyze_planning.py")
    recorded = report_path()
    with tempfile.TemporaryDirectory(prefix="litchi-0530-analysis-", dir="/home/zhuhe") as folder:
        output = Path(folder) / "replay.json"
        attempts = ([sys.executable, "-B", str(ANALYZER), str(output)],
                    [sys.executable, "-B", str(ANALYZER), "--output", str(output)])
        ok = False
        for command in attempts:
            result = subprocess.run(command, cwd=R, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode == 0 and output.is_file():
                ok = True
                break
            output.unlink(missing_ok=True)
        req(ok, "analyze_planning.py replay did not complete")
        req(output.read_bytes() == recorded.read_bytes(), "planning analyzer replay differs")
    value = read(recorded)
    req(isinstance(value, dict) and value.get("status") == "pass"
        and value.get("plan_sha256") == digest(PLAN)
        and value.get("owner") in OWNERS and isinstance(value.get("rows"), list)
        and len(value["rows"]) == 4, "planning report binding differs")
    text = json.dumps(value, sort_keys=True).lower()
    req("no latency" in text and "production change" in text,
        "planning report omits no-speedup scope")
    for key in ("numerical_speedup_claim", "runtime_speedup_claim", "speedup_claim"):
        if key in value:
            req(value[key] in (False, None, "none", "no"), f"planning report claims speedup: {key}")
    return {"report": str(recorded.relative_to(H)), "report_sha256": digest(recorded),
            "exact_replay": True, "performance_claim": "none"}


def validate_cleanup() -> dict[str, Any]:
    value = read(H / "cleanup.json")
    req(value.get("plan_sha256") == digest(PLAN) and value.get("removed") == OWNED
        and value.get("accessible_process_references") == []
        and value.get("owned_paths_absent") is True and value.get("python_cache_absent") is True,
        "cleanup receipt differs")
    req(all(not Path(path).exists() for path in OWNED) and not list(H.rglob("__pycache__")),
        "owned tree or Python cache remains")
    return {"owned_paths_absent": True, "python_cache_absent": True}


def check(name: str) -> dict[str, Any]:
    return {
        "source": validate_source,
        "prior": validate_prior,
        "build": lambda: {"receipt": validate_build()[0]["path"]},
        "builds": lambda: {"receipt": validate_build()[0]["path"]},
        "profiles": validate_profiles,
        "analysis": validate_analysis,
        "context": validate_context,
        "draft": validate_draft,
        "cleanup": validate_cleanup,
        "seal": lambda: {"entries": validate_seal(H / "SHA256SUMS", H)},
    }[name]()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=("all", "source", "prior", "build", "builds",
                                                 "profiles", "analysis", "context", "draft", "cleanup", "seal"),
                        default="all")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--output", type=Path, help="optional output outside this bundle")
    args = parser.parse_args()
    names = ("source", "prior", "build", "profiles", "analysis", "context", "draft", "cleanup", "seal") \
        if args.component == "all" else (args.component,)
    results, failed, pending = {}, False, False
    for name in names:
        try:
            results[name] = {"status": "pass", "result": check(name)}
        except Pending as error:
            pending, results[name] = True, {"status": "pending", "reason": str(error)}
        except EvidenceError as error:
            failed, results[name] = True, {"status": "fail", "reason": str(error)}
    status = "fail" if failed else "incomplete" if pending else "pass"
    output = {"schema": "litchi-0530-attribution-verification-v1", "status": status,
              "performance_claim": "none", "components": results}
    text = json.dumps(output, indent=2, sort_keys=True) + "\n"
    if args.output:
        target, root = args.output.resolve(), H.resolve()
        req(target != root and root not in target.parents, "refusing output inside this bundle")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
    print(text, end="")
    return 1 if failed or (pending and args.strict) else 0


if __name__ == "__main__":
    raise SystemExit(main())
