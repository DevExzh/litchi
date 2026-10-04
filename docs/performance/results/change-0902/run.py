"""Source-bound capture driver for the paired 0902 XML-audit experiment.

The driver deliberately contains no analysis or acceptance decision.  It runs
only the frozen child schedule described by ``plan.json`` and records enough
custody information for a later, independent reducer.  A child is never
retried: stdout, stderr, the JSON report, resource observation, and a receipt
are retained even when the child exits unsuccessfully.

The root coordinator writes ``plan.json`` after both roles have been built.
This module is kept packet-local so a later experiment cannot silently reuse
the 0902 schedule or source/binary bindings.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Any


PACKET = Path(__file__).resolve().parent
REPO = PACKET.parents[3]
PLAN_PATH = PACKET / "plan.json"
SCHEMA = "litchi.paired-audit-capture-0902.v1"
ROLES = ("baseline", "candidate")
ROLE_KINDS = ("native", "allocator")
SAFE_NAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")


class CaptureError(RuntimeError):
    """A frozen-capture precondition or child-capture failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error


def canonical_json(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        raise CaptureError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def resolve_path(value: str | Path, base: Path = PACKET) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base / path).absolute()


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite retained artifact {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
    try:
        with path.open("x", encoding="utf-8") as output:
            output.write(encoded)
    except FileExistsError as error:
        raise CaptureError(f"refusing to overwrite retained artifact {path}") from error


def file_descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"artifact must not be a symlink {path}")
    path = path.resolve()
    require(path.exists(), f"missing artifact {path}")
    require(path.is_file(), f"artifact is not a regular file {path}")
    stat = path.stat()
    return {
        "path": str(path),
        "bytes": stat.st_size,
        "mode": stat.st_mode & 0o777,
        "executable": bool(stat.st_mode & 0o111),
        "sha256": sha256_file(path),
    }


def descriptor_path(value: Any, label: str) -> Path:
    if isinstance(value, str):
        return resolve_path(value)
    require(isinstance(value, dict), f"{label} must be a path or descriptor")
    require(isinstance(value.get("path"), str), f"{label}.path missing")
    return resolve_path(value["path"])


def verify_descriptor(value: Any, label: str) -> dict[str, Any]:
    """Verify a plan-owned descriptor without trusting its recorded metadata."""

    path = descriptor_path(value, label)
    actual = file_descriptor(path)
    if isinstance(value, dict):
        for key in ("bytes", "sha256"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} does not match {path}")
        if "mode" in value:
            require(actual["mode"] == value["mode"], f"{label}.mode does not match {path}")
        if "executable" in value:
            require(actual["executable"] == value["executable"],
                    f"{label}.executable does not match {path}")
    return actual


def load_plan() -> dict[str, Any]:
    require(PLAN_PATH.is_file(), f"missing frozen plan {PLAN_PATH}")
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan must be an object")
    schema = plan.get("schema")
    require(schema == SCHEMA, f"unexpected plan schema {schema!r}")
    status = plan.get("status")
    require(status in {"frozen", "capture-ready", "ready"},
            f"plan is not frozen for capture: {status!r}")
    require(plan.get("plan_id") == "0902", "plan_id must be 0902")
    require(plan.get("seed") == 902009, "plan seed must be 902009")
    gate = plan.get("gate")
    require(isinstance(gate, dict), "plan.gate must be frozen")
    bootstrap = gate.get("bootstrap")
    primary = gate.get("primary")
    review = gate.get("review")
    require(isinstance(bootstrap, dict) and bootstrap.get("replicates") == 10000 and
            bootstrap.get("seed") == 902009 and
            bootstrap.get("statistic") == "paired_median_ratio",
            "gate.bootstrap must be 10,000 paired-median-ratio replicates with seed 902009")
    require(isinstance(primary, dict) and primary.get("p50_median_ratio_max") == 0.98 and
            primary.get("upper_95_ci_lt") == 1.0,
            "gate.primary must require p50 median <=0.98 and upper 95% CI <1")
    require(isinstance(review, dict) and review.get("regression_ratio") == 1.05 and
            review.get("p50_spread") == 1.20 and
            review.get("tail_metrics") == ["p50", "p95", "p99"],
            "gate.review regression thresholds are not frozen")
    return plan


def environment(plan: dict[str, Any]) -> dict[str, str]:
    """Return the complete child environment from the flat frozen mapping."""

    value = plan.get("environment")
    if isinstance(value, dict) and "fixed" in value:
        value = value["fixed"]
    require(isinstance(value, dict), "plan.environment must be a flat mapping")
    result: dict[str, str] = {}
    for key, item in value.items():
        require(isinstance(key, str) and key and re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", key),
                f"invalid environment key {key!r}")
        require(isinstance(item, str), f"environment value {key} must be a string")
        result[key] = item
    for key in ("PATH", "LC_ALL", "LANG", "TZ"):
        require(key in result, f"environment is missing fixed {key}")
    return dict(sorted(result.items()))


def source_specs(plan: dict[str, Any]) -> tuple[Any, Any, set[str], Path, Path | None]:
    source = plan.get("source")
    require(isinstance(source, dict), "plan.source must be an object")
    baseline = source.get("baseline_manifest")
    candidate = source.get("candidate_manifest")
    require(baseline is not None and candidate is not None,
            "plan.source requires baseline_manifest and candidate_manifest")
    names = source.get("candidate_files")
    require(isinstance(names, list) and len(names) == 2,
            "exactly two candidate source files are required")
    candidate_files: set[str] = set()
    for name in names:
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"invalid candidate source path {name!r}")
        path = Path(name)
        require(".." not in path.parts and path.as_posix() == name,
                f"candidate source path is not normalized {name!r}")
        candidate_files.add(name)
    root_value = source.get("root_worktree", str(REPO))
    require(isinstance(root_value, str), "source.root_worktree must be a path")
    root = resolve_path(root_value, PACKET)
    source_root_value = source.get("source_root")
    source_root = None
    if source_root_value is not None:
        require(isinstance(source_root_value, str), "source.source_root must be a path")
        source_root = resolve_path(source_root_value, PACKET)
    return baseline, candidate, candidate_files, root, source_root


def manifest_files(value: Any, label: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = descriptor_path(value, label)
    verify_descriptor(value, label)
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    require(isinstance(files, dict) and files, f"{label} has no files map")
    normalized: dict[str, dict[str, Any]] = {}
    for name, item in files.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{label} has invalid path {name!r}")
        path_name = Path(name)
        require(".." not in path_name.parts and path_name.as_posix() == name,
                f"{label} has non-normalized path {name!r}")
        require(isinstance(item, dict) and isinstance(item.get("sha256"), str),
                f"{label} entry {name} lacks sha256")
        normalized[name] = item
    return path, normalized


def source_check(plan: dict[str, Any], full: bool = True) -> dict[str, Any]:
    """Check the current candidate tree and the exact two-file source delta.

    Full checking hashes every manifest entry at stage boundaries.  Per-child
    checks hash the two candidate paths, keeping the custody guard cheap while
    still detecting mutation of the files under test immediately.  The final
    full check catches any other source mutation before the stage is sealed.
    """

    baseline_spec, candidate_spec, candidate_names, root, source_root = source_specs(plan)
    require(root.is_dir(), f"source root does not exist: {root}")
    _, baseline = manifest_files(baseline_spec, "baseline source manifest")
    _, candidate = manifest_files(candidate_spec, "candidate source manifest")
    require(set(baseline) == set(candidate), "baseline/candidate source file sets differ")
    changed = {
        name for name in baseline
        if (baseline[name].get("sha256"), baseline[name].get("bytes"))
        != (candidate[name].get("sha256"), candidate[name].get("bytes"))
    }
    require(changed == candidate_names,
            f"source delta is not exactly candidate_files: {sorted(changed)}")

    checked = 0
    paths = sorted(candidate_names if not full else candidate)
    for name in paths:
        path = root / name
        actual = file_descriptor(path)
        expected = candidate[name]
        require(actual["bytes"] == expected.get("bytes"), f"candidate source bytes differ: {name}")
        require(actual["sha256"] == expected["sha256"], f"candidate source differs: {name}")
        checked += 1

    source_root_state: dict[str, Any] = {"path": str(source_root) if source_root else None,
                                         "present": bool(source_root and source_root.is_dir()),
                                         "checked": False}
    if source_root is not None:
        require(source_root.is_dir(), f"isolated source root does not exist: {source_root}")
    if source_root is not None and full:
        for name, expected in candidate.items():
            actual = file_descriptor(source_root / name)
            require(actual["bytes"] == expected.get("bytes") and
                    actual["sha256"] == expected["sha256"],
                    f"isolated source root differs: {name}")
        source_root_state["checked"] = True
        source_root_state["file_count"] = len(candidate)

    return {
        "baseline_manifest": verify_descriptor(baseline_spec, "baseline source manifest"),
        "candidate_manifest": verify_descriptor(candidate_spec, "candidate source manifest"),
        "candidate_files": sorted(candidate_names),
        "checked_files": checked,
        "full": full,
        "root_worktree": str(root),
        "source_root": source_root_state,
    }


def role_descriptors(plan: dict[str, Any], role: str) -> dict[str, Any]:
    require(role in ROLES, f"unknown role {role}")
    roles = plan.get("roles", plan.get("binaries"))
    require(isinstance(roles, dict) and role in roles and isinstance(roles[role], dict),
            f"missing role descriptor {role}")
    result: dict[str, Any] = {}
    for kind in ROLE_KINDS:
        value = roles[role].get(kind)
        require(isinstance(value, dict), f"missing {role} {kind} descriptor")
        result[kind] = value
    return result


def candidate_input_bindings(raw: dict[str, Any], plan: dict[str, Any], label: str) -> None:
    before = raw.get("inputs_before")
    after = raw.get("inputs_after")
    require(isinstance(before, list) and before, f"{label}.inputs_before is empty")
    require(isinstance(after, list) and after, f"{label}.inputs_after is empty")
    require(before == after, f"{label} inputs_before/inputs_after differ")
    _, _, candidate_names, _, source_root = source_specs(plan)
    require(source_root is not None,
            f"{label} cannot bind candidate inputs without source.source_root")
    _, candidate = manifest_files(plan["source"]["candidate_manifest"],
                                  "candidate source manifest")
    indexed: dict[str, list[dict[str, Any]]] = {}
    for item in after:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} has malformed input descriptor")
        indexed.setdefault(str(Path(item["path"]).resolve()), []).append(item)
    for name in sorted(candidate_names):
        expected_path = str((source_root / name).resolve())
        matches = indexed.get(expected_path, [])
        require(len(matches) == 1, f"{label} does not bind exactly one input for {name}")
        item = matches[0]
        require(item.get("bytes") == candidate[name].get("bytes") and
                item.get("sha256") == candidate[name].get("sha256"),
                f"{label} candidate input differs from manifest: {name}")


def check_build_receipt(value: Any, label: str, role: str,
                        plan: dict[str, Any]) -> dict[str, Any]:
    actual = verify_descriptor(value, label)
    raw = read_json(Path(actual["path"]))
    require(isinstance(raw, dict), f"{label} is not an object")
    require(raw.get("exit_code") == 0, f"{label} does not record exit_code=0")
    if role == "baseline":
        require(raw.get("status") in {"pass", "passed", "success"},
                f"{label} does not record a successful baseline build")
    if role == "candidate":
        require(raw.get("inputs_unchanged") is True,
                f"{label} must record inputs_unchanged=true")
        candidate_input_bindings(raw, plan, label)
    return actual


def check_binary(plan: dict[str, Any], role: str, kind: str) -> dict[str, Any]:
    descriptor = role_descriptors(plan, role)[kind]
    actual = verify_descriptor(descriptor, f"{role} {kind} binary")
    require(actual["executable"], f"{role} {kind} binary is not executable")
    build = descriptor.get("build_receipt")
    require(build is not None, f"{role} {kind} binary lacks build_receipt")
    build_actual = check_build_receipt(build, f"{role} {kind} build receipt", role, plan)
    return {"binary": actual, "build_receipt": build_actual}


def check_binaries(plan: dict[str, Any]) -> dict[str, Any]:
    return {role: {kind: check_binary(plan, role, kind) for kind in ROLE_KINDS}
            for role in ROLES}


def case_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows = plan.get("cases")
    require(isinstance(rows, list) and len(rows) == 4, "plan must define four cases")
    require(len({row.get("id") for row in rows if isinstance(row, dict)}) == 4,
            "case ids must be unique")
    for row in rows:
        require(isinstance(row, dict), "case row must be an object")
        require(isinstance(row.get("id"), str) and SAFE_NAME.fullmatch(row["id"]),
                f"invalid case id {row.get('id')!r}")
        require(isinstance(row.get("case"), str), f"case {row['id']} lacks case name")

    def find(case: str, shape_key: str, shape: str) -> dict[str, Any]:
        matched = [row for row in rows if row.get("case") == case and
                   (row.get(shape_key) or row.get("shape")) == shape]
        require(len(matched) == 1, f"expected one {case}/{shape} case")
        row = matched[0]
        expected_kind = "primary" if "percent" in case else "control"
        require(row.get("kind", row.get("class")) == expected_kind,
                f"{row['id']} must be classified {expected_kind}")
        if shape_key == "semantic_shape":
            require(row.get("semantic_shape", row.get("shape")) == shape,
                    f"{row['id']} semantic shape mismatch")
        else:
            require(row.get("xlsx_cell_crud_shape", row.get("shape")) == shape,
                    f"{row['id']} XLSX shape mismatch")
        return row

    ordered = [
        find("xlsx_source_backed_cell_values_one_percent_edit_save",
             "xlsx_cell_crud_shape", "medium"),
        find("xlsx_source_backed_cell_values_one_percent_edit_save",
             "xlsx_cell_crud_shape", "dense-sparse"),
        find("xlsx_source_backed_cell_values_one_edit_save",
             "xlsx_cell_crud_shape", "medium"),
        find("docx_semantic_one_edit_save", "semantic_shape", "medium"),
    ]
    return ordered


def native_settings(plan: dict[str, Any]) -> tuple[int, int, int, int]:
    native = plan.get("native")
    qualification = plan.get("qualification")
    require(isinstance(native, dict) and isinstance(qualification, dict),
            "native and qualification settings are required")
    blocks = native.get("blocks")
    samples = native.get("samples")
    warmup = native.get("warmup")
    q_samples = qualification.get("samples")
    q_warmup = qualification.get("warmup")
    require((blocks, samples, warmup, q_samples, q_warmup) == (6, 100, 10, 2, 0),
            "0902 native counts must be 6 blocks/100 samples/10 warmup and 2/0 qualification")
    return blocks, samples, warmup, q_samples


def allocator_settings(plan: dict[str, Any]) -> tuple[int, int, int]:
    value = plan.get("allocator")
    require(isinstance(value, dict), "allocator settings are required")
    repeats, samples, warmup = value.get("repeats"), value.get("samples"), value.get("warmup")
    require((repeats, samples, warmup) == (2, 3, 2),
            "0902 allocator counts must be 2 repeats/3 samples/2 warmup")
    return repeats, samples, warmup


def cpu(plan: dict[str, Any]) -> int:
    value = plan.get("cpu", plan.get("native", {}).get("cpu"))
    require(isinstance(value, int) and value >= 0, "plan.cpu must be a nonnegative integer")
    return value


def workload_cwd(plan: dict[str, Any]) -> Path:
    value = plan.get("workload_cwd", plan.get("source", {}).get("root_worktree", str(REPO)))
    require(isinstance(value, str), "workload_cwd must be a path")
    path = resolve_path(value, PACKET)
    require(path == REPO.resolve(), f"workload cwd must be repository root {REPO}")
    require(path.is_dir(), f"workload cwd does not exist {path}")
    return path


def command_for(plan: dict[str, Any], role: str, kind: str, row: dict[str, Any],
                samples: int, warmup: int, folder: Path) -> list[str]:
    descriptor = role_descriptors(plan, role)[kind]
    binary = resolve_path(descriptor["path"])
    command: list[str] = ["/usr/bin/taskset", "-c", str(cpu(plan))]
    resource = plan.get("resource", {})
    if isinstance(resource, dict) and resource.get("enabled", True):
        time_path = resource.get("time_path", "/usr/bin/time")
        require(isinstance(time_path, str) and time_path, "resource.time_path must be a string")
        resource_format = resource.get(
            "format",
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        )
        require(isinstance(resource_format, str), "resource.format must be a string")
        command += [time_path, "-f", resource_format, "-o", str(folder / "resource.json")]
    command += [str(binary), "--warmup", str(warmup), "--samples", str(samples),
                "--case", row["case"]]
    if row["case"].startswith("xlsx_"):
        shape = row.get("xlsx_cell_crud_shape", row.get("shape"))
        require(isinstance(shape, str), f"{row['id']} lacks xlsx_cell_crud_shape")
        command += ["--xlsx-cell-crud-shape", shape]
    elif row["case"].startswith("docx_"):
        shape = row.get("semantic_shape", row.get("shape"))
        require(isinstance(shape, str), f"{row['id']} lacks semantic_shape")
        command += ["--semantic-shape", shape]
    else:
        raise CaptureError(f"unsupported 0902 case {row['case']}")
    command += ["--json", str(folder / "report.json")]
    return command


def expected_jobs(plan: dict[str, Any], stage: str) -> list[dict[str, Any]]:
    cases = case_rows(plan)
    jobs: list[dict[str, Any]] = []
    if stage == "qualification":
        for role in ROLES:
            for row in cases:
                jobs.append({"stage": stage, "index": 0, "role": role,
                             "case": row["id"], "row": row})
    elif stage == "native":
        blocks, _, _, _ = native_settings(plan)
        for block in range(blocks):
            roles = ROLES if block % 2 == 0 else tuple(reversed(ROLES))
            block_cases = cases if block % 2 == 0 else list(reversed(cases))
            for row in block_cases:
                for role in roles:
                    jobs.append({"stage": stage, "index": block, "role": role,
                                 "case": row["id"], "row": row})
    elif stage == "allocator":
        repeats, _, _ = allocator_settings(plan)
        for repeat in range(repeats):
            for role in ROLES:
                for row in cases:
                    jobs.append({"stage": stage, "index": repeat, "role": role,
                                 "case": row["id"], "row": row})
    else:
        raise CaptureError(f"unknown capture stage {stage}")
    return jobs


def child_name(job: dict[str, Any]) -> str:
    stage = job["stage"]
    prefix = {"qualification": "q", "native": "b", "allocator": "r"}[stage]
    return f"{prefix}{job['index']:02d}-{job['role']}-{job['case']}"


def stage_counts(plan: dict[str, Any], stage: str) -> tuple[int, int, int]:
    jobs = expected_jobs(plan, stage)
    if stage == "qualification":
        samples, warmup = 2, 0
    elif stage == "native":
        _, samples, warmup, _ = native_settings(plan)
    else:
        _, samples, warmup = allocator_settings(plan)
    return len(jobs), samples, warmup


def resource_enabled(plan: dict[str, Any]) -> bool:
    value = plan.get("resource", {})
    return not isinstance(value, dict) or value.get("enabled", True)


def artifact_descriptors(folder: Path) -> list[dict[str, Any]]:
    return [file_descriptor(path) for path in sorted(folder.iterdir())
            if path.is_file() and path.name not in {"receipt.json"}]


def run_child(plan: dict[str, Any], job: dict[str, Any], binaries: dict[str, Any],
              source_start: dict[str, Any], env: dict[str, str], cwd: Path) -> dict[str, Any]:
    stage = job["stage"]
    name = child_name(job)
    folder = PACKET / stage / name
    require(not folder.exists(), f"capture output already exists; refusing retry {folder}")
    folder.mkdir(parents=True)
    role = job["role"]
    kind = "allocator" if stage == "allocator" else "native"
    descriptor = role_descriptors(plan, role)[kind]
    binary = binaries[role][kind]["binary"]
    require(binary == verify_descriptor(descriptor, f"{role} {kind} binary"),
            f"{role} {kind} binary changed before {name}")
    if stage == "qualification":
        samples, warmup = 2, 0
    elif stage == "native":
        _, samples, warmup, _ = native_settings(plan)
    else:
        _, samples, warmup = allocator_settings(plan)
    command = command_for(plan, role, kind, job["row"], samples, warmup, folder)
    start = utc_now()
    start_record = {
        "schema": SCHEMA,
        "stage": stage,
        "index": job["index"],
        "role": role,
        "case_id": job["case"],
        "case": job["row"],
        "argv": command,
        "cwd": str(cwd),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "plan": verify_descriptor(PLAN_PATH, "plan"),
        "helper": file_descriptor(Path(__file__)),
        "binary": binary,
        "source": source_start,
        "started_utc": start,
    }
    write_once(folder / "start.json", start_record)

    exit_code: int | None = None
    exception: str | None = None
    tick = time.monotonic()
    try:
        with (folder / "stdout.txt").open("x", encoding="utf-8") as stdout, \
                (folder / "stderr.txt").open("x", encoding="utf-8") as stderr:
            result = subprocess.run(command, cwd=cwd, env=env, stdout=stdout, stderr=stderr,
                                    check=False)
            exit_code = result.returncode
    except (OSError, subprocess.SubprocessError) as error:
        exception = repr(error)
    elapsed = time.monotonic() - tick

    source_end: dict[str, Any] | None = None
    source_error: str | None = None
    try:
        source_end = source_check(plan, full=False)
    except CaptureError as error:
        source_error = str(error)

    failures: list[str] = []
    if exception:
        failures.append(f"subprocess exception: {exception}")
    if exit_code != 0:
        failures.append(f"child exit code: {exit_code}")
    if source_error:
        failures.append(f"source changed during child: {source_error}")
    if exit_code == 0 and not (folder / "report.json").is_file():
        failures.append("child exited zero without report.json")
    if resource_enabled(plan) and exit_code == 0 and not (folder / "resource.json").is_file():
        failures.append("child exited zero without resource.json")
    if failures:
        write_once(folder / "validation-failure.json", {
            "schema": SCHEMA,
            "status": "fail",
            "errors": failures,
            "retained_stdout": (folder / "stdout.txt").is_file(),
            "retained_stderr": (folder / "stderr.txt").is_file(),
        })

    command_record = {
        **start_record,
        "finished_utc": utc_now(),
        "elapsed_seconds": elapsed,
        "exit_code": exit_code,
        "exception": exception,
        "source_end": source_end,
        "source_error": source_error,
        "artifacts": artifact_descriptors(folder),
    }
    write_once(folder / "command.json", command_record)
    receipt = {
        "schema": SCHEMA,
        "status": "pass" if not failures else "fail",
        "stage": stage,
        "index": job["index"],
        "role": role,
        "case_id": job["case"],
        "argv": command,
        "cwd": str(cwd),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "plan": verify_descriptor(PLAN_PATH, "plan"),
        "helper": file_descriptor(Path(__file__)),
        "binary": binary,
        "source_start": source_start,
        "source_end": source_end,
        "source_error": source_error,
        "started_utc": start,
        "finished_utc": command_record["finished_utc"],
        "elapsed_seconds": elapsed,
        "exit_code": exit_code,
        "exception": exception,
        "errors": failures,
        "artifacts": artifact_descriptors(folder),
    }
    write_once(folder / "receipt.json", receipt)
    require(not failures, f"capture child failed ({name}); retained receipt and stderr")
    return receipt


def stage_capture(plan: dict[str, Any], stage: str) -> None:
    if stage == "native":
        qualification_capture = PACKET / "qualification" / "capture.json"
        require(qualification_capture.is_file(), "native capture requires completed qualification")
        q = read_json(qualification_capture)
        require(q.get("status") == "pass", "qualification is not sealed")
    source_start = source_check(plan, full=True)
    binaries = check_binaries(plan)
    env = environment(plan)
    cwd = workload_cwd(plan)
    jobs = expected_jobs(plan, stage)
    expected_children, samples, warmup = stage_counts(plan, stage)
    require(len(jobs) == expected_children, "internal schedule count mismatch")
    receipts = []
    for job in jobs:
        receipts.append(run_child(plan, job, binaries, source_start, env, cwd))
    source_end = source_check(plan, full=True)
    write_once(PACKET / stage / "capture.json", {
        "schema": SCHEMA,
        "status": "pass",
        "stage": stage,
        "children": len(receipts),
        "samples_per_child": samples,
        "warmup_per_child": warmup,
        "total_samples": len(receipts) * samples,
        "jobs": [{"name": child_name(job), "stage": job["stage"], "index": job["index"],
                  "role": job["role"], "case_id": job["case"]} for job in jobs],
        "plan": verify_descriptor(PLAN_PATH, "plan"),
        "helper": file_descriptor(Path(__file__)),
        "environment": env,
        "environment_sha256": sha256_bytes(canonical_json(env)),
        "source_start": source_start,
        "source_end": source_end,
        "binary_bindings": binaries,
        "receipts": receipts,
    })
    print(f"{stage}: {len(receipts)} children / {len(receipts) * samples} samples passed", flush=True)


def check_only(plan: dict[str, Any]) -> None:
    rows = case_rows(plan)
    native_settings(plan)
    allocator_settings(plan)
    source = source_check(plan, full=True)
    binaries = check_binaries(plan)
    env = environment(plan)
    cwd = workload_cwd(plan)
    expected = {stage: len(expected_jobs(plan, stage))
                for stage in ("qualification", "native", "allocator")}
    print(json.dumps({"status": "pass", "schema": SCHEMA, "cases": [row["id"] for row in rows],
                     "expected_children": expected, "cwd": str(cwd), "environment": env,
                     "source": source, "binaries": binaries}, indent=2, sort_keys=True))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["check", "qualification", "native", "allocator", "all"])
    args = parser.parse_args()
    plan = load_plan()
    if args.action == "check":
        check_only(plan)
        return
    if args.action == "all":
        stage_capture(plan, "qualification")
        stage_capture(plan, "native")
        stage_capture(plan, "allocator")
        return
    stage_capture(plan, args.action)


if __name__ == "__main__":
    try:
        main()
    except CaptureError as error:
        print(f"0902 capture failed: {error}", file=sys.stderr)
        raise SystemExit(1)
