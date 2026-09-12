"""Read-only verifier for the 0541 XLSX test evidence bundle.

The batch has no benchmark or production lane.  Every source attempt is a
separate snapshot, and only the final snapshot is required to be the current
source.  A failed focused test is retained as evidence; the six full quality
checks on the designated final snapshot must pass.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
from pathlib import Path
import re
import subprocess
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
QUALITY_PLAN = HERE / "quality-plan.json"
ADR = HERE / "adr-manifest.json"
FROZEN = HERE / "frozen-inputs.json"
DECISION = HERE / "decision.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
TARGET = Path("/home/zhuhe/litchi-goal-0541-target")

SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
FULL_QUALITY = (
    "quality-tests",
    "quality-check",
    "quality-clippy",
    "quality-rustdoc",
    "quality-fmt",
    "quality-boundaries",
)
FOCUSED = "focused-tests"
QUALITY_NAMES = FULL_QUALITY + (FOCUSED,)
ALLOWED_ADR_ROOT = "docs/adr/"
TEST_RESULT = re.compile(
    r"test result:\s+(ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;"
    r"\s+(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out"
)


class EvidenceError(ValueError):
    """Malformed or contradictory retained evidence."""


class Pending(EvidenceError):
    """A source, receipt, cleanup record, or seal has not arrived yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence bundle: {path}") from error


def display_path(path: Path) -> str:
    try:
        return relative(path)
    except EvidenceError:
        return path.as_posix()


def need(path: Path, label: str | None = None, *, directory: bool = False) -> Path:
    label = label or relative(path)
    if not path.exists():
        raise Pending(f"{label} is missing")
    if path.is_symlink():
        raise EvidenceError(f"{label} is a symlink")
    if directory:
        require(path.is_dir(), f"{label} is not a directory")
    else:
        require(path.is_file(), f"{label} is not a regular file")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or relative(path)
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {label}: {error}") from error


def read_bytes(path: Path, label: str | None = None) -> bytes:
    label = label or relative(path)
    need(path, label)
    try:
        return path.read_bytes()
    except OSError as error:
        raise EvidenceError(f"cannot read {label}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    try:
        return read_bytes(path, label).decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label or relative(path)} is not UTF-8") from error


def sha(path: Path) -> str:
    need(path, f"artifact for hashing: {display_path(path)}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise EvidenceError(f"cannot hash {display_path(path)}: {error}") from error
    return digest.hexdigest()


def valid_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and ".." not in path.parts and "." not in path.parts,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is invalid") from error
    require(result.tzinfo is not None, f"{label} has no timezone")
    return result


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds > 0 and end > start,
            f"{label} interval is invalid")
    return start, end


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def git_output(args: list[str], *, input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, input=input_data, stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise EvidenceError(
            f"Git command failed ({' '.join(args)}): {detail.decode(errors='replace')[-2000:]}"
        ) from error


def source_inventory() -> set[str]:
    tracked = git_output([
        "git", "ls-files", "-z", "--", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ]).split(b"\0")
    untracked = git_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ]).split(b"\0")
    names = {item.decode() for item in tracked if item}
    names.update(item.decode() for item in untracked if item and item.endswith(b".rs"))
    return {
        name for name in names
        if source_name(name) and (REPO / name).is_file() and not (REPO / name).is_symlink()
    }


def revision_manifest(revision: str) -> dict[str, str]:
    raw = git_output([
        "git", "ls-tree", "-r", "-z", revision, "--", "crates", "tools/perf-baseline",
        "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ])
    entries: list[tuple[str, bytes]] = []
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            metadata, encoded_name = item.split(b"\t", 1)
            fields = metadata.split()
            require(len(fields) == 3 and fields[1] == b"blob", "Git source tree entry is malformed")
            name = encoded_name.decode("utf-8")
        except (UnicodeDecodeError, ValueError) as error:
            raise EvidenceError("Git source tree entry is malformed") from error
        require(source_name(name), f"Git source tree contains out-of-scope source: {name}")
        entries.append((name, fields[2]))
    if not entries:
        raise EvidenceError("Git source tree is empty")
    request = b"".join(oid + b"\n" for _, oid in entries)
    raw = git_output(["git", "cat-file", "--batch"], input_data=request)
    result: dict[str, str] = {}
    position = 0
    for name, expected_oid in entries:
        end = raw.find(b"\n", position)
        require(end >= 0, "Git object response is truncated")
        fields = raw[position:end].split()
        require(len(fields) == 3 and fields[0] == expected_oid and fields[1] == b"blob",
                "Git source object response is malformed")
        position = end + 1
        length = int(fields[2])
        data = raw[position:position + length]
        require(len(data) == length, "Git source object is truncated")
        result[name] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n", "Git source object separator is missing")
        position += 1
    require(position == len(raw), "Git source object response has trailing data")
    require(len(result) == len(entries), "Git source tree repeats a path")
    return result


def manifest(path: Path) -> dict[str, str]:
    value = read_json(path, f"{relative(path)} source manifest")
    require(isinstance(value, dict) and value, f"{relative(path)} is not a nonempty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{relative(path)} source path")
        require(source_name(name) and valid_digest(digest),
                f"{relative(path)} has an invalid source entry: {name}")
        require(name not in result, f"{relative(path)} repeats {name}")
        result[name] = digest
    return result


def plan_data() -> dict[str, Any]:
    value = read_json(PLAN)
    require(isinstance(value, dict), "plan is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError("plan revision is not a Git commit") from error
    allowed = value.get("allowed_changes")
    require(isinstance(allowed, list) and {
        "crates/litchi-xlsx/tests/source_backed_cell_values.rs",
        "crates/litchi-xlsx/tests/source_backed_cell_values/planning_error_order.rs",
    } == set(allowed), "plan allowed changes differ")
    require(value.get("owned_paths") == [str(TARGET)], "plan owned paths differ")
    priority = value.get("priority")
    require(isinstance(priority, str) and "OLE2/OOXML" in priority and "ODF" in priority,
            "plan priority differs")
    require(isinstance(value.get("scope"), str)
            and "no production optimization" in value["scope"].lower()
            and "performance claim" in value["scope"].lower(),
            "plan scope is not test-only")
    return value


def quality_commands() -> dict[str, list[str]]:
    value = read_json(QUALITY_PLAN)
    require(isinstance(value, dict) and isinstance(value.get("commands"), dict),
            "quality plan is not an object")
    commands = value["commands"]
    require(set(commands) == set(QUALITY_NAMES), "quality command inventory differs")
    result: dict[str, list[str]] = {}
    for name in QUALITY_NAMES:
        command = commands[name]
        require(isinstance(command, list) and command and all(isinstance(item, str) for item in command),
                f"{name} command is malformed")
        lowered_tokens = {item.lower() for item in command}
        require(not lowered_tokens.intersection({"build", "profile", "perf", "callgrind"})
                and not any(item.lower().startswith(("--profile", "--callgrind"))
                            for item in command),
                f"{name} introduces a build or profile lane")
        result[name] = command
    require("cargo" in commands["quality-tests"] and "test" in commands["quality-tests"],
            "full test command is not a cargo test")
    require("--test" in commands[FOCUSED]
            and "source_backed_cell_values" in commands[FOCUSED]
            and "planning_error_order" in commands[FOCUSED],
            "focused command does not select planning_error_order")
    return result


def frozen_inputs() -> dt.datetime:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            "frozen input envelope differs")
    expected = {"plan.json", "run.py", "quality-plan.json", "adr-manifest.json"}
    require(set(value["files"]) == expected, "frozen input inventory differs")
    for name, digest in value["files"].items():
        require(valid_digest(digest) and sha(HERE / name) == digest,
                f"frozen input hash differs: {name}")
    return parse_time(value.get("created_utc"), "frozen-inputs.created_utc")


def adr_manifest() -> dict[str, str]:
    value = read_json(ADR, "adr-manifest.json")
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            "ADR manifest is malformed")
    files = value["files"]
    expected = set(
        item.decode() for item in git_output(["git", "ls-files", "-z", "--", "docs/adr"])
        .split(b"\0") if item
    )
    require(set(files) == expected, "ADR manifest inventory differs")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith(ALLOWED_ADR_ROOT) and valid_digest(digest),
                f"ADR entry is malformed: {name}")
        require((REPO / name).is_file() and not (REPO / name).is_symlink()
                and sha(REPO / name) == digest, f"ADR hash differs: {name}")
    return files


def stage_names() -> list[str]:
    names: list[str] = []
    for path in sorted(HERE.iterdir(), key=lambda item: item.name):
        if not path.is_dir() or path.is_symlink():
            continue
        if (path / "source-manifest.json").exists():
            safe_relative(path.name, "stage name")
            require("/" not in path.name, "stage name is nested")
            names.append(path.name)
    require(names, "no source attempt has been frozen")
    return names


def changed_paths(snapshot: dict[str, str], revision: dict[str, str]) -> set[str]:
    return {
        name for name in set(snapshot) | set(revision)
        if snapshot.get(name) != revision.get(name)
    }


def patch_paths(data: bytes, label: str) -> set[str]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label} is not UTF-8") from error
    result: set[str] = set()
    for line in text.splitlines():
        if not line.startswith("diff --git "):
            continue
        fields = line.split()
        require(len(fields) >= 4, f"{label} has a malformed diff header")
        for field in fields[2:4]:
            if field in ("/dev/null", "a//dev/null", "b//dev/null"):
                continue
            if field.startswith("a/") or field.startswith("b/"):
                field = field[2:]
            safe_relative(field, f"{label} path")
            result.add(field)
    return result


def git_show(revision: str, name: str) -> bytes:
    return git_output(["git", "show", f"{revision}:{name}"])


def validate_patch(stage: Path, snapshot: dict[str, str], revision: dict[str, str],
                  allowed: set[str], revision_name: str) -> set[str]:
    patch_path = need(stage / "source.patch", f"{relative(stage)}/source.patch")
    data = read_bytes(patch_path)
    changed = changed_paths(snapshot, revision)
    require(changed <= allowed, f"{relative(stage)} changes outside allowed test files: "
            f"{sorted(changed - allowed)}")
    paths = patch_paths(data, relative(patch_path))
    require(paths == changed, f"{relative(stage)} source patch paths differ")
    if not changed:
        require(data == b"", f"{relative(stage)} has a patch without source changes")
        return changed

    # Replay the patch against just the changed files from the frozen Git
    # revision.  This keeps the verifier read-only with respect to the repo,
    # while checking that the patch really produces the retained source bytes.
    # /tmp is intentionally not used: concurrent workspace campaigns can
    # exhaust its quota even though this replay only needs a few source files.
    with tempfile.TemporaryDirectory(prefix="litchi-0541-verify-", dir="/dev/shm") as temporary:
        root = Path(temporary)
        for name in changed & set(revision):
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(git_show(revision_name, name))
        try:
            checked = subprocess.run(["git", "apply", "--check", "--whitespace=nowarn", "-"],
                                     cwd=root, input=data, stdout=subprocess.PIPE,
                                     stderr=subprocess.PIPE, check=False)
            require(checked.returncode == 0,
                    f"{relative(patch_path)} does not apply: "
                    f"{checked.stderr.decode(errors='replace')[-1200:]}")
            applied = subprocess.run(["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
                                     input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                     check=False)
            require(applied.returncode == 0,
                    f"{relative(patch_path)} could not be replayed: "
                    f"{applied.stderr.decode(errors='replace')[-1200:]}")
        except OSError as error:
            raise EvidenceError(f"cannot replay {relative(patch_path)}: {error}") from error
        source_root = stage / "sources"
        for name in changed:
            produced = root / name
            retained = source_root / name
            if name in snapshot:
                require(produced.is_file() and retained.is_file()
                        and produced.read_bytes() == retained.read_bytes(),
                        f"{relative(stage)} patch output differs for {name}")
            else:
                require(not produced.exists() and not retained.exists(),
                        f"{relative(stage)} deletion custody differs for {name}")
    return changed


def validate_sources(stage_name: str, revision: dict[str, str], allowed: set[str]) -> dict[str, Any]:
    stage = HERE / stage_name
    require(stage.is_dir() and not stage.is_symlink(), f"{stage_name} is not a regular stage directory")
    snapshot = manifest(stage / "source-manifest.json")
    changed = validate_patch(stage, snapshot, revision, allowed, plan_data()["revision"])
    require(set(snapshot) - set(revision) <= allowed,
            f"{stage_name} adds an out-of-scope source")
    require(set(revision) - set(snapshot) <= allowed,
            f"{stage_name} removes an out-of-scope source")
    for name in set(snapshot) & set(revision):
        if name not in allowed:
            require(snapshot[name] == revision[name], f"{stage_name} changes {name}")

    sources = need(stage / "sources", f"{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"{stage_name}/sources path")
            actual.add(name)
    expected = set(snapshot) & allowed
    require(actual == expected, f"{stage_name}/sources inventory differs")
    for name in actual:
        copy = sources / name
        require(sha(copy) == snapshot[name], f"{stage_name}/sources hash differs: {name}")
    return {"manifest": snapshot, "changed": changed, "manifest_sha256": sha(stage / "source-manifest.json")}


def test_counts(stdout: str, label: str, *, allow_empty: bool = False) -> dict[str, int]:
    rows = list(TEST_RESULT.finditer(stdout))
    if not rows and allow_empty and "test result:" not in stdout:
        # A focused compile can fail before the harness starts and therefore
        # has no cargo summary.  Retain that attempt with a zero test count.
        return {"passed": 0, "failed": 0, "ignored": 0,
                "executed": 0, "result_lines": 0}
    require(rows, f"{label} has no cargo test result")
    passed = failed = ignored = 0
    for row in rows:
        passed += int(row.group(2))
        failed += int(row.group(3))
        ignored += int(row.group(4))
    return {"passed": passed, "failed": failed, "ignored": ignored,
            "executed": passed + failed, "result_lines": len(rows)}


def artifact_map(stage: Path, action: str, value: dict[str, Any]) -> None:
    entries = value.get("artifacts")
    require(isinstance(entries, dict), f"{relative(stage)}/{action} artifacts are missing")
    expected = {f"{action}.stdout", f"{action}.stderr"}
    actual = {
        path.name for path in stage.glob(action + ".*")
        if path.is_file() and not path.is_symlink() and path.name != f"{action}.receipt.json"
    }
    require(set(entries) == actual and expected <= actual,
            f"{relative(stage)}/{action} artifact inventory differs")
    for name, digest in entries.items():
        safe_relative(name, f"{relative(stage)}/{action} artifact")
        require(Path(name).name == name and valid_digest(digest),
                f"{relative(stage)}/{action} artifact entry is malformed")
        target = stage / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{relative(stage)}/{action} artifact hash differs: {name}")


def validate_receipt(stage_name: str, action: str, commands: dict[str, list[str]],
                     source_sha: str) -> dict[str, Any]:
    stage = HERE / stage_name
    path = stage / f"{action}.receipt.json"
    value = read_json(path, relative(path))
    require(isinstance(value, dict), f"{relative(path)} is not an object")
    require(value.get("command") == commands[action], f"{relative(path)} command differs")
    exit_code = value.get("exit_code")
    require(isinstance(exit_code, int) and not isinstance(exit_code, bool),
            f"{relative(path)} exit code is malformed")
    start, end = interval(value, relative(path))
    require(value.get("source_manifest_sha256") == source_sha
            and value.get("script_sha256") == sha(RUN)
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("quality_plan_sha256") == sha(QUALITY_PLAN)
            and value.get("source_unchanged") is True,
            f"{relative(path)} source or driver binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict) and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{relative(path)} temporary directory binding differs")
    artifact_map(stage, action, value)
    counts: dict[str, int] | None = None
    if action in ("quality-tests", FOCUSED):
        stdout = read_text(stage / f"{action}.stdout", relative(stage / f"{action}.stdout"))
        counts = test_counts(
            stdout, relative(stage / f"{action}.stdout"),
            # A failed quality command may stop during compilation and emit
            # no Cargo summary.  Keep that earlier attempt auditable; the
            # designated final stage is checked for successful summaries
            # below.
            allow_empty=exit_code != 0,
        )
    return {"name": action, "path": relative(path), "value": value,
            "start": start, "end": end, "sha256": sha(path), "counts": counts}


def validate_stages(names: list[str], commands: dict[str, list[str]], revision: dict[str, str],
                    allowed: set[str], frozen: dt.datetime) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    sources: dict[str, dict[str, Any]] = {}
    for stage_name in names:
        sources[stage_name] = validate_sources(stage_name, revision, allowed)
        receipts = sorted((HERE / stage_name).glob("*.receipt.json"), key=lambda item: item.name)
        if not receipts:
            raise Pending(f"{stage_name} has no quality receipts")
        actual_actions = {path.name.removesuffix(".receipt.json") for path in receipts}
        require(actual_actions <= set(QUALITY_NAMES),
                f"{stage_name} contains a build or profile receipt")
        for action in sorted(actual_actions):
            record = validate_receipt(stage_name, action, commands,
                                      sources[stage_name]["manifest_sha256"])
            require(record["start"] > frozen, f"{record['path']} predates frozen inputs")
            records.append(record)
    ordered = sorted(records, key=lambda item: item["start"])
    require(all(left["end"] <= right["start"] for left, right in zip(ordered, ordered[1:])),
            "quality receipt intervals overlap")
    by_stage = {
        name: [record for record in records if record["path"].startswith(name + "/")]
        for name in names
    }
    return {"records": records, "sources": sources, "by_stage": by_stage,
            "receipt_count": len(records)}


def validate_final(names: list[str], stage_data: dict[str, Any], commands: dict[str, list[str]]) -> dict[str, Any]:
    decision = read_json(DECISION, "decision.json")
    require(isinstance(decision, dict), "decision is not an object")
    final = decision.get("final_stage")
    safe_relative(final, "decision.final_stage")
    require("/" not in final and final in names, "decision final stage is not a source attempt")
    focused_count = decision.get("focused_test_count")
    full_count = decision.get("full_test_count")
    require(isinstance(focused_count, int) and not isinstance(focused_count, bool) and focused_count >= 0,
            "decision focused test count is malformed")
    require(isinstance(full_count, int) and not isinstance(full_count, bool) and full_count >= 0,
            "decision full test count is malformed")

    final_records = stage_data["by_stage"][final]
    final_actions = {record["name"] for record in final_records}
    require(final_actions == set(QUALITY_NAMES),
            "final stage does not contain exactly six full and one focused quality receipt")
    require(all(record["value"].get("exit_code") == 0 for record in final_records),
            "final stage contains a failed quality receipt")
    final_focused = next(record for record in final_records if record["name"] == FOCUSED)
    require(final_focused["counts"] is not None,
            "final focused test has no Cargo summary")
    require(focused_count == final_focused["counts"]["passed"],
            "decision focused test count differs from final stdout")
    final_tests = next(record for record in final_records if record["name"] == "quality-tests")
    require(final_tests["counts"] is not None
            and full_count == final_tests["counts"]["passed"],
            "decision full test count differs from stdout")
    focused = [record for record in stage_data["records"] if record["name"] == FOCUSED]
    prior_focused = sum(
        record["counts"]["passed"] for record in stage_data["records"]
        if record["name"] == FOCUSED and record["path"].split("/", 1)[0] != final
        and record["counts"] is not None
    )
    return {"final_stage": final, "focused_test_count": focused_count,
            "full_test_count": full_count, "focused_receipts": len(focused),
            "prior_focused_test_count": prior_focused,
            "full_quality_receipts": len(FULL_QUALITY)}


def validate_current_source(final_stage: str, stage_data: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    final_manifest = stage_data["sources"][final_stage]["manifest"]
    current_names = source_inventory()
    require(set(final_manifest) == current_names, "current source inventory differs from final manifest")
    for name, digest in final_manifest.items():
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"current source hash differs: {name}")
    revision = revision_manifest(plan["revision"])
    changed = changed_paths(final_manifest, revision)
    require(changed <= set(plan["allowed_changes"]), "current source changes outside plan")
    return {"manifest_sha256": sha(HERE / final_stage / "source-manifest.json"),
            "entries": len(final_manifest), "changed_files": sorted(changed)}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    if not CLEANUP.exists():
        raise Pending("cleanup.json is missing")
    value = read_json(CLEANUP, "cleanup.json")
    require(value.get("removed") == plan["owned_paths"]
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == [],
            "cleanup record differs")
    for owned in plan["owned_paths"]:
        target = Path(owned)
        require(not target.exists() and not target.is_symlink(), f"owned target remains: {owned}")
    return {"owned_paths_absent": True, "removed": plan["owned_paths"]}


def snapshot() -> dict[str, str]:
    result: dict[str, str] = {}
    for path in HERE.rglob("*"):
        require(not path.is_symlink(), f"evidence bundle contains symlink: {relative(path)}")
        if path.is_file() and path != SEAL:
            result[relative(path)] = sha(path)
    return result


def validate_seal() -> dict[str, Any]:
    need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(SEAL, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "malformed SHA256SUMS line")
        name = safe_relative(fields[1], "SHA256SUMS path")
        require(name != "SHA256SUMS" and name not in expected, "SHA256SUMS repeats or seals itself")
        expected[name] = fields[0]
    actual = snapshot()
    require(actual == expected, "SHA256SUMS inventory or digest differs")
    return {"entries": len(expected), "sha256": sha(SEAL)}


def verify(*, precleanup: bool = False) -> dict[str, Any]:
    before = snapshot()
    plan = plan_data()
    commands = quality_commands()
    frozen = frozen_inputs()
    adr_manifest()
    revision = revision_manifest(plan["revision"])
    names = stage_names()
    allowed = set(plan["allowed_changes"])
    stages = validate_stages(names, commands, revision, allowed, frozen)
    decision = validate_final(names, stages, commands)
    source = validate_current_source(decision["final_stage"], stages, plan)
    if CLEANUP.exists():
        cleanup: dict[str, Any] = {"status": "pass", **validate_cleanup(plan)}
    elif SEAL.exists():
        raise EvidenceError("SHA256SUMS exists before cleanup.json")
    elif precleanup:
        cleanup = {"status": "pending", "reason": "cleanup.json is pending"}
    else:
        raise Pending("cleanup.json is missing")
    if SEAL.exists():
        seal: dict[str, Any] = {"status": "pass", **validate_seal()}
    elif precleanup:
        seal = {"status": "pending", "reason": "SHA256SUMS is pending"}
    else:
        raise Pending("SHA256SUMS is missing")
    after = snapshot()
    require(before == after, "verifier mutated evidence bundle")
    status = "pass" if cleanup["status"] == seal["status"] == "pass" else "incomplete"
    return {
        "status": status,
        "scope": "Test-only XLSX planning error-order evidence; no performance claim",
        "final_stage": decision["final_stage"],
        "stages": names,
        "receipts": stages["receipt_count"],
        "focused_test_count": decision["focused_test_count"],
        "full_test_count": decision["full_test_count"],
        "prior_focused_test_count": decision["prior_focused_test_count"],
        "source": source,
        "adr_entries": len(adr_manifest()),
        "cleanup": cleanup,
        "seal": seal,
        "benchmark_builds": 0,
        "profiles": 0,
        "performance_claim": "none",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true",
                        help="allow cleanup.json and SHA256SUMS to be pending")
    parser.add_argument("--strict", action="store_true",
                        help="require cleanup.json and SHA256SUMS (the default)")
    parser.add_argument("--output", type=Path, help="write compact JSON outside this bundle")
    args = parser.parse_args()
    precleanup = args.precleanup and not args.strict
    before = snapshot()
    try:
        result = verify(precleanup=precleanup)
        output = result
    except Pending as error:
        output = {"status": "incomplete", "reason": str(error),
                  "scope": "Test-only XLSX planning error-order evidence; no performance claim"}
    except EvidenceError as error:
        output = {"status": "fail", "reason": str(error),
                  "scope": "Test-only XLSX planning error-order evidence; no performance claim"}
    after = snapshot()
    if before != after:
        output = {"status": "fail", "reason": "verifier mutated evidence bundle",
                  "scope": "Test-only XLSX planning error-order evidence; no performance claim"}
    text = json.dumps(output, indent=2, sort_keys=True) + "\n"
    if args.output:
        target = args.output.resolve()
        require(target != HERE.resolve() and HERE.resolve() not in target.parents,
                "refusing output inside evidence bundle")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
    print(text, end="")
    return 1 if output["status"] == "fail" or (output["status"] == "incomplete" and not precleanup) else 0


if __name__ == "__main__":
    raise SystemExit(main())
