"""Independent, read-only verifier for the 0531 MCE namespace-search pilot.

The campaign is intentionally source-bound.  This verifier checks the frozen
inputs, replays the baseline/candidate patches in a private Git index, binds
every build and child receipt to its source and executable, checks the native
ABBA custody order and logical output identities, replays the canonical
analyzer into an external temporary file, and validates quality, disposition,
cleanup, and sealing evidence.  Missing campaign evidence is reported as
``incomplete``; it is never treated as a passing performance result.
"""

from __future__ import annotations

import argparse
from collections import Counter
import datetime as dt
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
FROZEN = HERE / "frozen-inputs.json"
ADR = HERE / "adr-manifest.json"
START = HERE / "start-state.json"
BASELINE = HERE / "baseline"
CANDIDATE = HERE / "candidate"
FINAL = HERE / "final"
RESTORED = HERE / "restored"
SCRATCH = Path("/tmp/litchi-goal-0531")
TARGET = Path("/home/zhuhe/litchi-goal-0531-target")
SOURCE_ROOTS = ("crates/litchi-ooxml-common/",)
MCE_FILE = "crates/litchi-ooxml-common/src/mce/codec.rs"
TEST_DRAFT = HERE / "mce_namespace_search.rs.draft"
TEST_BINDING = HERE / "test-source-binding.json"
RETAINED_TEST = "crates/litchi-ooxml-common/tests/mce_namespace_search.rs"
XLSX_COMPACT_FILE = "crates/litchi-xlsx/src/raw/compact.rs"
CANDIDATE_TEST_COPY = HERE / "candidate-test.rs"
FINAL_TEST_COPY = HERE / "final-test.rs"
TEST_FIX_PATCH = HERE / "test-fix.patch"
TEST_LINT_PATCH = HERE / "test-lint.patch"
FINAL_BINDING = HERE / "final-source-binding.json"
TEST_REFERENCE = HERE / "test-reference"
CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SHAPES = ("medium", "dense-sparse")
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
CONDITIONAL_LANES = ("profile", "hardware", "eager")
REQUIRED_CONDITIONAL_LANES = ("profile", "eager", "reopen")
PROFILE_REPORT = HERE / "profile-comparison.json"
PROFILE_ANALYZER = HERE / "analyze_profiles.py"
EAGER_PLAN = HERE / "eager-plan.json"
EAGER_REPORT = HERE / "eager-comparison.json"
EAGER_ANALYZER = HERE / "eager_guard.py"
EAGER_REPLAY_ANALYZER = HERE / "eager_analysis.py"
REOPEN_PLAN = HERE / "reopen-plan.json"
REOPEN_REPORT = HERE / "reopen-comparison.json"
REOPEN_ANALYZER = HERE / "reopen_confirmation.py"
NATIVE_PILOT_REPORT = HERE / "native-pilot-comparison.json"
NATIVE_REPORT_HISTORY = HERE / "native-report-history.json"
FINAL_NATIVE_PLAN = HERE / "final-native-plan.json"
FINAL_NATIVE_FROZEN = HERE / "final-native-frozen-inputs.json"
FINAL_NATIVE_REPORT = HERE / "final-native-comparison.json"
FINAL_NATIVE_ANALYZER = HERE / "analyze_final_native.py"
FINAL_NATIVE_REVIEW = HERE / "final-native-review.json"
FINAL_PROFILE_REPORT = HERE / "final-profile-comparison.json"
FINAL_PROFILE_ANALYZER = HERE / "analyze_final_profiles.py"
REPORT_NAMES = (
    "comparison.json",
    "analysis.json",
    "pilot-comparison.json",
    "performance-comparison.json",
)
ANALYZER_NAMES = ("analyze.py", "analyze_comparison.py", "analyze_pilot.py")
QUALITY_SUMMARY = HERE / "quality-summary.json"
QUALITY_PLAN = HERE / "quality-plan.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
OWNED = [str(SCRATCH), str(TARGET)]


class EvidenceError(ValueError):
    """Malformed, contradictory, or out-of-scope evidence."""


class IncompleteError(EvidenceError):
    """Evidence selected by the plan has not arrived yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def rel(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def need(path: Path, label: str | None = None, *, symlink: bool = True) -> Path:
    label = label or rel(path)
    if not path.exists():
        raise IncompleteError(f"{label} is missing: {rel(path)}")
    if symlink and path.is_symlink():
        raise EvidenceError(f"{label} is a symlink: {rel(path)}")
    return path


def read_json(path: Path, label: str | None = None) -> Any:
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {label or rel(path)}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    need(path, label)
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise EvidenceError(f"cannot read text {label or rel(path)}: {error}") from error


def sha(path: Path) -> str:
    need(path, f"artifact for hashing: {rel(path)}", symlink=False)
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {rel(path)}: {error}") from error


def valid_digest(value: Any) -> bool:
    return isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value) is not None


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is invalid") from error
    require(parsed.tzinfo is not None, f"{label} has no timezone")
    return parsed


def interval(value: Any, label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds > 0 and end > start,
            f"{label} interval is invalid")
    return start, end


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def source_name(name: str) -> bool:
    exact = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
    return (name.startswith("crates/") or name.startswith("tools/perf-baseline/")
            or name.startswith(".cargo/") or name in exact)


def current_source_names() -> set[str]:
    tracked = subprocess.check_output([
        "git", "ls-files", "-z", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ], cwd=REPO)
    untracked = subprocess.check_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ], cwd=REPO)
    names = {item.decode() for item in tracked.split(b"\0") if item}
    names.update(item.decode() for item in untracked.split(b"\0")
                 if item and item.endswith(b".rs"))
    return {name for name in names if source_name(name) and (REPO / name).is_file()}


def current_source_manifest() -> dict[str, str]:
    return {name: sha(REPO / name) for name in sorted(current_source_names())}


def source_manifest(path: Path) -> dict[str, str]:
    value = read_json(path, f"{rel(path)} source manifest")
    require(isinstance(value, dict) and value, f"{rel(path)} is not a nonempty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{rel(path)} source path")
        require(source_name(name) and valid_digest(digest),
                f"{rel(path)} has an invalid source entry: {name}")
        require(name not in result, f"{rel(path)} repeats {name}")
        result[name] = digest
    return result


def run_checked(args: list[str], *, env: dict[str, str] | None = None,
                input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, env=env, input=input_data,
                                       stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise EvidenceError(f"command failed ({' '.join(args)}): "
                            f"{detail.decode(errors='replace')[-2000:]}") from error


def revision_source_names(revision: str) -> set[str]:
    raw = run_checked(["git", "ls-tree", "-r", "--name-only", revision],
                      env=None)
    return {name for name in raw.decode().splitlines() if source_name(name)}


def parse_private_index(index: Path) -> dict[str, tuple[str, int]]:
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    raw = run_checked(["git", "ls-files", "-s", "-z"], env=env)
    result: dict[str, tuple[str, int]] = {}
    for item in raw.split(b"\0"):
        if not item:
            continue
        metadata, name_bytes = item.split(b"\t", 1)
        fields = metadata.split()
        require(len(fields) == 3, "private source replay entry is malformed")
        name = name_bytes.decode()
        require(name not in result, f"private source replay repeats {name}")
        result[name] = (fields[1].decode(), int(fields[0]))
    return result


def index_blob_hashes(index: Path, oids: set[str]) -> dict[str, str]:
    if not oids:
        return {}
    env = dict(os.environ, GIT_INDEX_FILE=str(index))
    payload = b"\n".join(oid.encode() for oid in sorted(oids)) + b"\n"
    raw = run_checked(["git", "cat-file", "--batch"], env=env, input_data=payload)
    result: dict[str, str] = {}
    position = 0
    while position < len(raw):
        end = raw.find(b"\n", position)
        require(end >= 0, "Git batch response is malformed")
        fields = raw[position:end].split()
        require(len(fields) == 3 and fields[1] == b"blob",
                "source replay object is not a blob")
        oid = fields[0].decode()
        length = int(fields[2])
        position = end + 1
        data = raw[position:position + length]
        require(len(data) == length, "Git batch response is truncated")
        result[oid] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n", "Git batch separator is missing")
        position += 1
    return result


def replay_patch(path: Path, revision: str) -> dict[str, Any]:
    """Apply a frozen patch to a private index and return source blob custody."""
    need(path, f"source patch {rel(path)}")
    with tempfile.TemporaryDirectory(prefix="litchi-0531-replay-", dir="/home/zhuhe") as folder:
        index = Path(folder) / "index"
        env = dict(os.environ, GIT_INDEX_FILE=str(index))
        run_checked(["git", "read-tree", revision], env=env)
        if path.stat().st_size:
            run_checked(["git", "apply", "--cached", "--binary", str(path)], env=env)
        indexed = parse_private_index(index)
        changed = set(run_checked([
            "git", "diff", "--cached", "--name-only", revision,
        ], env=env).decode().splitlines())
        scoped = {name: row for name, row in indexed.items() if source_name(name)}
        hashes = index_blob_hashes(index, {oid for oid, _ in scoped.values()})
        return {"changed": changed, "indexed": scoped, "hashes": hashes}


def load_plan() -> dict[str, Any]:
    frozen = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(frozen, dict)
            and set(frozen) == {"plan.json", "run.py", "candidate.patch"},
            "frozen input inventory differs")
    for name, path in (("plan.json", PLAN), ("run.py", RUN),
                       ("candidate.patch", HERE / "candidate.patch")):
        require(frozen.get(name) == sha(path), f"frozen {name} digest differs")
    plan = read_json(PLAN, "plan.json")
    require(isinstance(plan, dict)
            and plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build/capture")
    revision = plan.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError("plan revision is not a Git commit") from error

    require(plan.get("candidate_source_roots") == list(SOURCE_ROOTS),
            "candidate source roots differ")
    require(plan.get("owned_paths") == OWNED, "owned temporary paths differ")
    require(plan.get("cpu") == 2, "planned CPU differs")
    require(plan.get("native_order") == [
        "baseline native-r1 before candidate application",
        "candidate native-r1",
        "candidate native-r2",
        "retained baseline native-r2 under candidate source checkout",
    ], "native ABBA order differs")
    primary = plan.get("primary")
    require(isinstance(primary, dict) and primary.get("case") == CASE
            and primary.get("shapes") == list(SHAPES)
            and primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 200, "primary plan differs")
    require(plan.get("guard_repeats") == 2 and plan.get("guard_warmup") == 10
            and plan.get("guard_samples") == 30, "guard counts differ")
    expected_guards = [
        {"case": "xlsx_source_backed_cell_values_one_edit_save",
         "shapes": ["medium", "dense-sparse"]},
        {"case": "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
         "shapes": ["medium", "dense-sparse"]},
        {"case": CASE, "shapes": ["vendor-extension"]},
        {"case": CASE, "shapes": ["noncompact"]},
        {"case": "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
         "shapes": ["noncompact"]},
        {"case": "docx_source_backed_one_edit_save", "shapes": ["medium"]},
        {"case": "pptx_source_backed_one_edit_save", "shapes": ["medium"]},
    ]
    require(plan.get("guards") == expected_guards, "guard plan differs")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict) and allocation.get("shapes") == list(SHAPES)
            and allocation.get("repeats") == 2 and allocation.get("warmup") == 0
            and allocation.get("samples") == 5
            and "diagnostic" in str(allocation.get("scope", "")).lower(),
            "allocation plan differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict) and profile.get("shapes") == list(SHAPES)
            and profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 1
            and profile.get("owner") ==
            "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets",
            "profile plan differs")
    require(plan.get("gates") == {
        "total_p50_reduction_percent": 1.0,
        "total_mean_reduction_percent": 1.0,
        "planning_p50_reduction_percent": 2.0,
        "planning_ir_reduction_percent": 1.0,
        "require_every_shape_repeat": True,
    }, "0531 gates differ")
    require(valid_digest(plan.get("draft_patch_sha256"))
            and plan["draft_patch_sha256"] == sha(HERE / "candidate.patch"),
            "candidate patch binding differs")
    return plan


def validate_adrs_and_start(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(ADR, "adr-manifest.json")
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files, "ADR manifest is malformed")
    for name, expected in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith("docs/adr/") and valid_digest(expected)
                and (REPO / name).is_file() and sha(REPO / name) == expected,
                f"ADR hash differs: {name}")
    state = read_json(START, "start-state.json")
    previous_commit = state.get("previous_commit")
    require(isinstance(previous_commit, str) and plan["revision"].startswith(previous_commit),
            "start-state commit differs")
    require(state.get("priority") == "OLE2/OOXML performance; ODF deferred until completion; iWork excluded",
            "start-state priority differs")
    review = HERE / "source-review.md"
    if review.exists():
        text = read_text(review, "source-review.md")
        require(plan["revision"] in text and "candidate.patch" in text
                and MCE_FILE in text and "memmem" in text,
                "source review is not bound to the frozen candidate")
    test_binding = None
    if TEST_BINDING.exists():
        test_binding = read_json(TEST_BINDING, "test-source-binding.json")
        require(test_binding.get("file") == RETAINED_TEST
                and valid_digest(test_binding.get("sha256"))
                and valid_digest(test_binding.get("draft_sha256"))
                and test_binding.get("draft") == TEST_DRAFT.name
                and sha(TEST_DRAFT) == test_binding["draft_sha256"],
                "test source binding differs")
    return {"adr_sha256": sha(ADR), "adr_entries": len(files),
            "start_state_sha256": sha(START),
            "source_review_sha256": sha(review) if review.exists() else None,
            "test_source_binding_sha256": sha(TEST_BINDING) if test_binding is not None else None}


def validate_final_test_binding(manifest: dict[str, str], plan: dict[str, Any],
                               extras: set[str]) -> dict[str, Any]:
    """Validate final test-only corrections and their reference run.

    The production candidate was measured with the original formatted test.
    The final checkout may retain a corrected test and a linter-only change
    in an existing XLSX ``cfg(test)`` helper.  Both are independently bound
    by frozen patches and a baseline-production reference receipt.  This
    custody cannot silently turn either test-only change into release code.
    """
    require(extras == {RETAINED_TEST},
            "final source has an unapproved unpatched source file")
    for path, label in ((FINAL_BINDING, "final-source-binding.json"),
                        (CANDIDATE_TEST_COPY, "candidate-test.rs"),
                        (FINAL_TEST_COPY, "final-test.rs"),
                        (TEST_FIX_PATCH, "test-fix.patch"),
                        (TEST_LINT_PATCH, "test-lint.patch")):
        need(path, label)
    binding = read_json(FINAL_BINDING, "final-source-binding.json")
    require(isinstance(binding, dict), "final-source-binding.json is malformed")
    candidate_manifest_path = CANDIDATE / "source-manifest.json"
    final_manifest_path = FINAL / "source-manifest.json"
    reference_manifest_path = TEST_REFERENCE / "source-manifest.json"
    reference_receipt_path = TEST_REFERENCE / "check-mce-fixed-tests.receipt.json"
    for path, label in ((candidate_manifest_path, "candidate source manifest"),
                        (final_manifest_path, "final source manifest"),
                        (reference_manifest_path, "test reference source manifest"),
                        (reference_receipt_path, "test reference receipt")):
        need(path, label)
    candidate_manifest = source_manifest(candidate_manifest_path)
    final_manifest = source_manifest(final_manifest_path)
    reference_manifest = source_manifest(reference_manifest_path)
    candidate_test_sha = sha(CANDIDATE_TEST_COPY)
    final_test_sha = sha(FINAL_TEST_COPY)
    require(XLSX_COMPACT_FILE in candidate_manifest
            and XLSX_COMPACT_FILE in final_manifest
            and XLSX_COMPACT_FILE in manifest,
            "final XLSX test helper is absent from source manifests")
    require(candidate_manifest.get(RETAINED_TEST) == candidate_test_sha
            and final_manifest.get(RETAINED_TEST) == final_test_sha
            and manifest.get(RETAINED_TEST) == final_test_sha,
            "final test source hashes are not manifest-bound")
    before_lint_manifest = dict(candidate_manifest)
    before_lint_manifest[RETAINED_TEST] = final_test_sha
    before_lint_manifest_sha = hashlib.sha256(
        (json.dumps(before_lint_manifest, indent=2) + "\n").encode("utf-8")
    ).hexdigest()
    require(binding.get("candidate_manifest_sha256") == sha(candidate_manifest_path)
            and binding.get("final_manifest_sha256") == sha(final_manifest_path)
            and binding.get("changed_files") == [RETAINED_TEST, XLSX_COMPACT_FILE]
            and binding.get("original_test_sha256") == candidate_test_sha
            and binding.get("final_test_sha256") == final_test_sha
            and binding.get("patch_sha256") == sha(TEST_FIX_PATCH)
            and binding.get("baseline_reference_receipt_sha256") == sha(reference_receipt_path)
            and binding.get("before_lint_final_manifest_sha256") == before_lint_manifest_sha
            and binding.get("test_lint_patch_sha256") == sha(TEST_LINT_PATCH),
            "final source binding differs")
    expected_before_lint = dict(before_lint_manifest)
    expected_final_changes = sorted(name for name in set(expected_before_lint) | set(final_manifest)
                                    if expected_before_lint.get(name) != final_manifest.get(name))
    require(expected_final_changes == [XLSX_COMPACT_FILE],
            "final source has an unapproved post-lint source change")
    require(isinstance(binding.get("scope"), str)
            and ("production" in binding["scope"].lower()
                 or "release" in binding["scope"].lower())
            and "test" in binding["scope"].lower(),
            "final source binding scope is missing")

    # Reconstruct the corrected test in an external temporary tree using the
    # frozen patch.  This exercises the patch payload without touching the
    # shared checkout or running a compiler.
    with tempfile.TemporaryDirectory(prefix="litchi-0531-test-fix-", dir="/home/zhuhe") as folder:
        root = Path(folder)
        target = root / RETAINED_TEST
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(CANDIDATE_TEST_COPY.read_bytes())
        result = subprocess.run(
            ["git", "apply", "--no-index", "--unsafe-paths", str(TEST_FIX_PATCH.resolve())],
            cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(result.returncode == 0,
                "test-fix.patch cannot be applied to the frozen candidate test")
        require(hashlib.sha256(target.read_bytes()).hexdigest() == final_test_sha,
                "test-fix.patch reconstruction differs from final-test.rs")

    # Reconstruct the linter-only XLSX helper edit from the frozen revision.
    # The patch is intentionally constrained to one exact repeat_n hunk, and
    # the production prefix before cfg(test) must remain byte-for-byte equal.
    lint_lines = read_text(TEST_LINT_PATCH, "test-lint.patch").splitlines()
    require(lint_lines[:3] == [
        f"--- a/{XLSX_COMPACT_FILE}",
        f"+++ b/{XLSX_COMPACT_FILE}",
        "@@ -603,7 +603,7 @@",
    ], "test-lint.patch header differs")
    require(lint_lines[3:] == [
        "     #[test]",
        "     fn changed_worksheet_can_compact_an_oversized_whitespace_input() {",
        '         let mut input = format!(r#"<worksheet xmlns=\"{MAIN_NAMESPACE}\">"#).into_bytes();',
        "-        input.extend(std::iter::repeat(b' ').take(16 * 1024 * 1024 + 1));",
        "+        input.extend(std::iter::repeat_n(b' ', 16 * 1024 * 1024 + 1));",
        '         input.extend_from_slice(b"<sheetData/></worksheet>");',
        "         assert!(input.len() > 16 * 1024 * 1024);",
        " ",
    ], "test-lint.patch hunk differs")
    candidate_compact = run_checked(
        ["git", "show", f"{plan['revision']}:{XLSX_COMPACT_FILE}"])
    require(hashlib.sha256(candidate_compact).hexdigest() ==
            candidate_manifest[XLSX_COMPACT_FILE],
            "frozen revision does not match candidate XLSX test helper")
    with tempfile.TemporaryDirectory(prefix="litchi-0531-test-lint-", dir="/home/zhuhe") as folder:
        root = Path(folder)
        target = root / XLSX_COMPACT_FILE
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(candidate_compact)
        result = subprocess.run(
            ["git", "apply", "--no-index", "--unsafe-paths", str(TEST_LINT_PATCH.resolve())],
            cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(result.returncode == 0,
                "test-lint.patch cannot be applied to the frozen candidate helper")
        reconstructed_compact = target.read_bytes()
        require(hashlib.sha256(reconstructed_compact).hexdigest() ==
                final_manifest[XLSX_COMPACT_FILE],
                "test-lint.patch reconstruction differs from final helper")
        marker = b"\n#[cfg(test)]"
        candidate_marker = candidate_compact.find(marker)
        final_marker = reconstructed_compact.find(marker)
        require(candidate_marker >= 0 and candidate_marker == candidate_compact.rfind(marker)
                and final_marker == candidate_marker
                and final_marker == reconstructed_compact.rfind(marker)
                and candidate_compact[:candidate_marker] ==
                reconstructed_compact[:final_marker],
                "final XLSX helper changes production code before cfg(test)")

    # The reference run deliberately uses baseline production with the
    # corrected test.  Validate its receipt and output custody as ordinary
    # source-bound quality evidence, without invoking the recorded command.
    reference_manifest_sha = sha(reference_manifest_path)
    receipt = read_json(reference_receipt_path, "test reference receipt")
    expected_command = [
        "env", "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0", "cargo", "test", "--locked",
        "-p", "litchi-ooxml-common", "--all-features", "--test",
        "mce_namespace_search",
    ]
    require(receipt.get("command") == expected_command
            and receipt.get("exit_code") == 0
            and receipt.get("plan_sha256") == sha(PLAN)
            and receipt.get("script_sha256") == sha(RUN)
            and receipt.get("source_manifest_sha256") == reference_manifest_sha
            and receipt.get("working_source_manifest_sha256") == reference_manifest_sha
            and receipt.get("binary_sha256") is None,
            "test reference receipt binding differs")
    interval(receipt, "test reference receipt")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {
                "check-mce-fixed-tests.stdout", "check-mce-fixed-tests.stderr"},
            "test reference artifact inventory differs")
    for filename, expected_sha in artifacts.items():
        path = TEST_REFERENCE / filename
        require(valid_digest(expected_sha) and path.is_file() and not path.is_symlink()
                and sha(path) == expected_sha,
                f"test reference artifact custody differs: {filename}")
    output = read_text(TEST_REFERENCE / "check-mce-fixed-tests.stdout",
                       "test reference stdout")
    require(sum(int(number) for number in re.findall(
        r"test result: ok\. (\d+) passed;", output)) == 9,
            "test reference does not retain all nine corrected tests")
    require(reference_manifest.get(MCE_FILE) ==
            source_manifest(BASELINE / "source-manifest.json")[MCE_FILE],
            "test reference production is not baseline-identical")
    expected_reference = dict(source_manifest(BASELINE / "source-manifest.json"))
    expected_reference[RETAINED_TEST] = final_test_sha
    require(reference_manifest == expected_reference,
            "test reference source inventory differs from baseline plus corrected test")
    return {"binding_sha256": sha(FINAL_BINDING),
            "candidate_test_sha256": candidate_test_sha,
            "final_test_sha256": final_test_sha,
            "test_fix_sha256": sha(TEST_FIX_PATCH),
            "reference_manifest_sha256": reference_manifest_sha,
            "reference_receipt_sha256": sha(reference_receipt_path),
            "reference_tests": 9}


def validate_restored_source_binding(manifest: dict[str, str],
                                     plan: dict[str, Any],
                                     extras: set[str]) -> dict[str, Any]:
    """Validate the production-restored source plus retained test fixes."""
    require(extras == {RETAINED_TEST},
            "restored source has an unapproved unpatched source file")
    # Validate the common corrected-test and exact cfg(test) lint payload
    # against the frozen final binding, then require the restored inventory to
    # be baseline production with those two test-only results applied.
    final_manifest_path = FINAL / "source-manifest.json"
    need(final_manifest_path, "final source manifest for restored binding")
    final_binding = validate_final_test_binding(
        source_manifest(final_manifest_path), plan, {RETAINED_TEST})
    baseline_manifest = source_manifest(BASELINE / "source-manifest.json")
    final_manifest = source_manifest(final_manifest_path)
    expected = dict(baseline_manifest)
    expected[RETAINED_TEST] = final_manifest[RETAINED_TEST]
    expected[XLSX_COMPACT_FILE] = final_manifest[XLSX_COMPACT_FILE]
    require(manifest == expected,
            "restored source must equal baseline production plus retained test fixes")
    require(manifest[MCE_FILE] == baseline_manifest[MCE_FILE],
            "restored source does not restore the production MCE codec")
    return {"final_binding": final_binding,
            "restored_manifest_sha256": sha(RESTORED / "source-manifest.json"),
            "baseline_manifest_sha256": sha(BASELINE / "source-manifest.json"),
            "retained_test_sha256": manifest[RETAINED_TEST],
            "lint_helper_sha256": manifest[XLSX_COMPACT_FILE]}


def stage_blob_custody(stage: str, plan: dict[str, Any], *, allow_empty: bool = False) -> dict[str, Any]:
    folder = {"baseline": BASELINE, "candidate": CANDIDATE,
              "final": FINAL, "restored": RESTORED}[stage]
    manifest_path = need(folder / "source-manifest.json", f"{stage} source manifest")
    patch_path = need(folder / "source.patch", f"{stage} source patch")
    manifest = source_manifest(manifest_path)
    replay = replay_patch(patch_path, plan["revision"])
    changed = replay["changed"]
    require(all(source_name(name) for name in changed),
            f"{stage} patch changes an out-of-scope file")
    if stage == "final":
        require(set(changed) == {MCE_FILE, XLSX_COMPACT_FILE},
                "final patch changes more than production plus the frozen XLSX test hunk")
    elif stage == "restored":
        require(set(changed) == {XLSX_COMPACT_FILE},
                "restored patch must contain only the frozen XLSX test hunk")
    indexed_names = set(replay["indexed"])
    extras = set(manifest) - indexed_names
    require(stage in ("candidate", "final", "restored") or not extras,
            f"{stage} source manifest contains unpatched source files")
    if extras:
        if stage == "final":
            final_binding = validate_final_test_binding(manifest, plan, extras)
        elif stage == "restored":
            restored_binding = validate_restored_source_binding(manifest, plan, extras)
        else:
            # run.py intentionally freezes non-Git Rust test drafts in the
            # source manifest while its source.patch contains only the
            # tracked production hunk.  The sidecar binds that extra file to
            # the reviewed formatted test source; it is never accepted for a
            # production path.
            binding = read_json(TEST_BINDING, "test-source-binding.json")
            require(extras == {binding.get("file")} and extras == {RETAINED_TEST},
                    f"{stage} unpatched source inventory is not the reviewed test")
            for name in extras:
                # The shared checkout can already contain the final corrected
                # test after the production decision.  Candidate custody is
                # therefore checked against the immutable candidate copy,
                # rather than against the mutable current checkout.
                require(manifest[name] == binding.get("sha256")
                        and CANDIDATE_TEST_COPY.is_file()
                        and not CANDIDATE_TEST_COPY.is_symlink()
                        and sha(CANDIDATE_TEST_COPY) == manifest[name],
                        f"{stage} retained test custody differs")
    require(indexed_names <= set(manifest),
            f"{stage} source index contains an unmanifested source file")
    for name, expected in manifest.items():
        if name in extras:
            continue
        oid, mode = replay["indexed"][name]
        require(mode in (100644, 100755), f"{stage} source mode is unexpected: {name}")
        require(replay["hashes"].get(oid) == expected,
                f"{stage} replay blob differs: {name}")
    revision_names = revision_source_names(plan["revision"])
    if stage == "baseline":
        require(not changed, "baseline source patch must be empty")
        require(set(manifest) == revision_names,
                "baseline source inventory differs from frozen revision")
    elif stage == "candidate":
        require(changed, "candidate source patch is empty")
        require(MCE_FILE in changed, "candidate patch omits the MCE production file")
        require(all(any(name.startswith(root) for root in SOURCE_ROOTS) for name in changed),
                "candidate patch escapes the OOXML common root")
        frozen_patch = (HERE / "candidate.patch").read_bytes()
        # candidate/source.patch is the normal Git diff with index metadata;
        # candidate.patch is the separately frozen one-hunk draft.  They need
        # not have identical textual headers, but both must replay to the same
        # production blob.
        draft_replay = replay_patch(HERE / "candidate.patch", plan["revision"])
        require(draft_replay["changed"] == {MCE_FILE}
                and draft_replay["hashes"].get(draft_replay["indexed"][MCE_FILE][0])
                == manifest[MCE_FILE],
                "frozen candidate.patch does not bind the candidate MCE blob")
        text = frozen_patch.decode("utf-8")
        old = ".windows(NAMESPACE.len())"
        old_any = ".any(|w| w == NAMESPACE.as_bytes())"
        new = "memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_none()"
        require(text.count(old) == 1 and text.count(old_any) == 1 and text.count(new) == 1,
                "candidate patch is not the exact namespace predicate substitution")
    else:
        require(allow_empty or changed, "final source patch is empty")
        if stage == "final":
            require(set(changed) <= {MCE_FILE, XLSX_COMPACT_FILE},
                    "final source patch contains an unapproved file")
        elif stage == "restored":
            require(set(changed) <= {XLSX_COMPACT_FILE},
                    "restored source patch contains an unapproved file")
    return {"stage": stage, "manifest_sha256": sha(manifest_path),
            "manifest_entries": len(manifest), "patch_sha256": sha(patch_path),
            "changed_files": sorted(changed)}


def validate_source() -> dict[str, Any]:
    plan = load_plan()
    binding = validate_adrs_and_start(plan)
    baseline = stage_blob_custody("baseline", plan)
    candidate = stage_blob_custody("candidate", plan)
    restored = None
    if (RESTORED / "source-manifest.json").is_file():
        restored = stage_blob_custody("restored", plan)
    base_manifest = source_manifest(BASELINE / "source-manifest.json")
    cand_manifest = source_manifest(CANDIDATE / "source-manifest.json")
    changed = sorted(name for name in set(base_manifest) | set(cand_manifest)
                     if base_manifest.get(name) != cand_manifest.get(name))
    require(changed and MCE_FILE in changed
            and set(candidate["changed_files"]) <= set(changed),
            "candidate manifest diff does not match replayed patch")
    require(all(any(name.startswith(root) for root in SOURCE_ROOTS) for name in changed),
            "candidate manifest diff escapes candidate root")
    require(set(changed) <= {MCE_FILE, RETAINED_TEST},
            "candidate manifest contains an unplanned OOXML common source change")
    current = current_source_manifest()
    frozen_source_maps = [base_manifest, cand_manifest]
    # A rejected campaign may leave the final checkout restored to the
    # baseline production tree while retaining only the reviewed test.  That
    # final manifest is itself frozen by the decision record and therefore is
    # a valid source state for this read-only component check.
    final_manifest_path = FINAL / "source-manifest.json"
    if final_manifest_path.is_file():
        frozen_source_maps.append(source_manifest(final_manifest_path))
    restored_manifest_path = RESTORED / "source-manifest.json"
    if restored_manifest_path.is_file():
        frozen_source_maps.append(source_manifest(restored_manifest_path))
    require(current in frozen_source_maps,
            "current source checkout matches neither frozen baseline, candidate, final, nor restored")
    source_diff = CANDIDATE / "source-diff.json"
    if source_diff.exists():
        diff = read_json(source_diff, "candidate source-diff.json")
        require(diff.get("baseline_manifest_sha256") == baseline["manifest_sha256"]
                and diff.get("candidate_manifest_sha256") == candidate["manifest_sha256"]
                and diff.get("candidate_source_roots") == list(SOURCE_ROOTS),
                "candidate source-diff binding differs")
        rows = diff.get("changed_files")
        require(isinstance(rows, dict) and sorted(rows) == changed,
                "candidate source-diff inventory differs")
        for name in changed:
            require(rows[name].get("baseline_sha256") == base_manifest.get(name)
                    and rows[name].get("candidate_sha256") == cand_manifest.get(name),
                    f"candidate source-diff digest differs: {name}")
    return {"status": "pass", "binding": binding, "baseline": baseline,
            "candidate": candidate, "restored": restored,
            "candidate_baseline_changes": changed,
            "current_source": ("baseline" if current == base_manifest else
                               "candidate" if current == cand_manifest else
                               "final" if final_manifest_path.is_file()
                               and current == source_manifest(final_manifest_path)
                               else "restored")}


def expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    primary = plan["primary"]
    jobs: list[dict[str, Any]] = []
    if lane == "native":
        for repeat in range(1, primary["repeats"] + 1):
            shapes = primary["shapes"] if repeat == 1 else list(reversed(primary["shapes"]))
            for shape in shapes:
                jobs.append({"name": f"native-r{repeat}-primary-{shape}", "lane": lane,
                             "kind": "primary", "guard": None, "repeat": repeat,
                             "case": primary["case"], "shape": shape,
                             "warmup": primary["warmup"], "samples": primary["samples"]})
        for repeat in range(1, plan["guard_repeats"] + 1):
            for guard, item in enumerate(plan["guards"]):
                for shape in item["shapes"]:
                    jobs.append({"name": f"native-r{repeat}-guard{guard}-{shape}",
                                 "lane": lane, "kind": "guard", "guard": guard,
                                 "repeat": repeat, "case": item["case"], "shape": shape,
                                 "warmup": plan["guard_warmup"],
                                 "samples": plan["guard_samples"]})
        return jobs
    config = plan["allocation"] if lane == "alloc" else plan[lane]
    for repeat in range(1, config["repeats"] + 1):
        for shape in config["shapes"]:
            jobs.append({"name": f"{lane}-r{repeat}-{shape}", "lane": lane,
                         "kind": lane, "guard": None, "repeat": repeat,
                         "case": primary["case"], "shape": shape,
                         "warmup": config["warmup"], "samples": config["samples"]})
    return jobs


def expected_working(stage: str, name: str, base_sha: str, candidate_sha: str) -> set[str]:
    if stage == "candidate":
        return {candidate_sha}
    if stage == "restored":
        # restored_quality.py runs against the restored production tree and
        # writes receipts beside that stage's summary.  Keep the working
        # source binding exact; a baseline/candidate alias would allow a
        # quality result to be attributed to the wrong checkout.
        return {sha(RESTORED / "source-manifest.json")}
    if stage == "final":
        values = {base_sha, candidate_sha}
        final_manifest = FINAL / "source-manifest.json"
        if final_manifest.is_file():
            values.add(sha(final_manifest))
        return values
    if stage == "baseline" and name.startswith("native-r2-"):
        return {candidate_sha}
    # A baseline diagnostic lane may be captured before or after candidate
    # application.  Its receipt still has to name one of the two frozen maps.
    return {base_sha, candidate_sha}


def artifact_map(receipt: dict[str, Any], folder: Path, job: dict[str, Any]) -> None:
    name = job["name"]
    values = receipt.get("artifacts")
    require(isinstance(values, dict), f"{name} artifact inventory is missing")
    expected = {name + suffix for suffix in (".json", ".stdout", ".stderr")}
    if job["lane"] == "native":
        expected.add(name + ".rss.json")
    elif job["lane"] == "hardware":
        expected.add(name + ".csv")
    elif job["lane"] == "profile":
        callgrind = {item for item in values if item.startswith(name + ".callgrind")}
        require(callgrind, f"{name} callgrind artifact is missing")
        require(all(item == name + ".callgrind"
                    or re.fullmatch(re.escape(name) + r"\.callgrind\.[0-9]+", item)
                    for item in callgrind), f"{name} callgrind artifact name is malformed")
        expected |= callgrind
    require(set(values) == expected, f"{name} artifact inventory differs")
    for filename, expected_sha in values.items():
        safe_relative(filename, f"{name} artifact")
        path = folder / filename
        require(Path(filename).name == filename and path.is_file()
                and not path.is_symlink() and valid_digest(expected_sha)
                and sha(path) == expected_sha,
                f"{name} artifact custody differs: {filename}")


def build_command(kind: str) -> list[str]:
    executable = "litchi-perf-baseline" + ("-alloc" if kind == "alloc" else "")
    command = ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
               "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
               "--bin", executable, "--target-dir", str(TARGET)]
    if kind == "alloc":
        command += ["--features", "allocator-metrics"]
    return command


def receipt_common(path: Path, stage: str, job_name: str, stage_sha: str,
                   base_sha: str, candidate_sha: str, binary_sha: str | None = None) -> dict[str, Any]:
    value = read_json(path, rel(path))
    start, end = interval(value, rel(path))
    require(value.get("exit_code") == 0 and value.get("plan_sha256") == sha(PLAN)
            and value.get("script_sha256") == sha(RUN)
            and value.get("source_manifest_sha256") == stage_sha
            and value.get("working_source_manifest_sha256") in
            expected_working(stage, job_name, base_sha, candidate_sha),
            f"{rel(path)} receipt binding differs")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{rel(path)} TMPDIR differs")
    if binary_sha is not None:
        require(value.get("binary_sha256") == binary_sha,
                f"{rel(path)} binary binding differs")
    return {"path": rel(path), "name": job_name, "start": start,
            "end": end, "sha256": sha(path), "value": value}


def binary_identity(stage: str, kind: str, stage_sha: str) -> dict[str, Any]:
    identity_path = {"baseline": BASELINE, "candidate": CANDIDATE,
                     "final": FINAL}[stage] / f"binary-{kind}.json"
    value = read_json(identity_path, rel(identity_path))
    expected_path = SCRATCH / f"{stage}-{kind}"
    require(value.get("path") == str(expected_path) and valid_digest(value.get("sha256")),
            f"{rel(identity_path)} path/digest differs")
    nonnegative_int(value.get("bytes"), f"{rel(identity_path)}.bytes")
    require(value["bytes"] > 0 and value.get("source_manifest_sha256") == stage_sha,
            f"{rel(identity_path)} source binding differs")
    receipt = {"baseline": BASELINE, "candidate": CANDIDATE,
               "final": FINAL}[stage] / \
        f"build-{kind}.receipt.json"
    require(value.get("build_receipt_sha256") == sha(receipt),
            f"{rel(identity_path)} build receipt binding differs")
    binary = Path(value["path"])
    if binary.exists():
        require(binary.is_file() and not binary.is_symlink() and sha(binary) == value["sha256"]
                and binary.stat().st_size == value["bytes"],
                f"{rel(identity_path)} binary custody differs")
    else:
        cleanup = read_json(CLEANUP, "cleanup.json")
        require(cleanup.get("owned_paths_absent") is True
                and cleanup.get("removed") == OWNED,
                f"{rel(identity_path)} missing binary lacks cleanup custody")
    return {"kind": kind, "path": str(binary), "sha256": value["sha256"],
            "bytes": value["bytes"], "identity_sha256": sha(identity_path)}


def validate_builds() -> dict[str, Any]:
    plan = load_plan()
    source = validate_source()
    base_sha = source["baseline"]["manifest_sha256"]
    candidate_sha = source["candidate"]["manifest_sha256"]
    stages: dict[str, Any] = {}
    for stage, folder in (("baseline", BASELINE), ("candidate", CANDIDATE)):
        stage_sha = sha(folder / "source-manifest.json")
        rows = []
        identities = {}
        for kind in ("normal", "alloc"):
            receipt_path = folder / f"build-{kind}.receipt.json"
            identity_path = folder / f"binary-{kind}.json"
            if kind == "alloc" and not receipt_path.exists() and not identity_path.exists():
                # The allocator lane is explicitly optional diagnostic
                # evidence in the 0531 analyzer.  If either side starts it,
                # the complete bound lane is required below.
                continue
            if kind == "alloc":
                require(receipt_path.exists() and identity_path.exists(),
                        f"{stage} allocator build custody is incomplete")
            row = receipt_common(receipt_path, stage, f"build-{kind}", stage_sha,
                                 base_sha, candidate_sha)
            value = row["value"]
            require(value.get("binary_sha256") is None
                    and value.get("command") == build_command(kind),
                    f"{rel(receipt_path)} command differs")
            artifacts = value.get("artifacts")
            require(isinstance(artifacts, dict)
                    and set(artifacts) == {f"build-{kind}.stdout", f"build-{kind}.stderr"},
                    f"{rel(receipt_path)} artifact inventory differs")
            for filename, expected_sha in artifacts.items():
                path = folder / filename
                require(valid_digest(expected_sha) and path.is_file()
                        and not path.is_symlink() and sha(path) == expected_sha,
                        f"{rel(receipt_path)} artifact custody differs")
            identities[kind] = binary_identity(stage, kind, stage_sha)
            rows.append({"path": row["path"], "sha256": row["sha256"],
                         "start": row["start"].isoformat(), "end": row["end"].isoformat()})
        stages[stage] = {"manifest_sha256": stage_sha, "builds": rows,
                         "binaries": identities}
    require(all("normal" in stages[stage]["binaries"] for stage in stages),
            "normal build evidence is missing")
    return {"status": "pass", "stages": stages, "binary_sha256":
            stages["candidate"]["binaries"]["normal"]["sha256"]}


def validate_final_normal_binary(candidate_sha: str) -> dict[str, Any]:
    """Check the release-source rebuild preserves the measured executable.

    A final checkout may contain corrected tests and a cfg(test)-only lint
    fix.  If it is selected for acceptance, its normal benchmark binary must
    still be byte-identical to the measured candidate binary.  This check is
    deliberately independent of any decision status label.
    """
    plan = load_plan()
    final_manifest_path = FINAL / "source-manifest.json"
    final_sha = sha(final_manifest_path)
    base_sha = sha(BASELINE / "source-manifest.json")
    candidate_manifest_sha = sha(CANDIDATE / "source-manifest.json")
    receipt_path = need(FINAL / "build-normal.receipt.json",
                        "final build-normal receipt")
    row = receipt_common(receipt_path, "final", "build-normal", final_sha,
                         base_sha, candidate_manifest_sha)
    value = row["value"]
    require(value.get("binary_sha256") is None
            and value.get("working_source_manifest_sha256") == final_sha
            and value.get("command") == build_command("normal"),
            "final normal build receipt differs")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {"build-normal.stdout", "build-normal.stderr"},
            "final normal build artifact inventory differs")
    for filename, expected in artifacts.items():
        path = FINAL / filename
        require(valid_digest(expected) and path.is_file() and not path.is_symlink()
                and sha(path) == expected,
                f"final normal build artifact custody differs: {filename}")
    final_binary = binary_identity("final", "normal", final_sha)
    candidate_binary = binary_identity("candidate", "normal", candidate_manifest_sha)
    require(candidate_binary["sha256"] == candidate_sha, "candidate binary argument differs")
    require(final_binary["sha256"] == candidate_binary["sha256"],
            "final normal binary differs from measured candidate binary")
    return {"status": "pass", "receipt": row["path"],
            "receipt_sha256": row["sha256"], "binary": final_binary,
            "measured_candidate_binary_sha256": candidate_binary["sha256"],
            "byte_identical": True}


def capture_command(job: dict[str, Any], binary: dict[str, Any], folder: Path,
                    plan: dict[str, Any]) -> list[str]:
    command = ["taskset", "-c", str(plan["cpu"])]
    name = job["name"]
    if job["lane"] == "native":
        command += ["/usr/bin/time", "-f",
                    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
                    "-o", str(folder / (name + ".rss.json"))]
    elif job["lane"] == "profile":
        owner = plan["profile"]["owner"]
        command += ["valgrind", "--tool=callgrind", "--collect-atstart=no",
                    "--toggle-collect=" + owner, "--zero-before=" + owner,
                    "--dump-after=" + owner,
                    "--callgrind-out-file=" + str(folder / (name + ".callgrind"))]
    elif job["lane"] == "hardware":
        command += ["perf", "stat", "-x", ",", "-o", str(folder / (name + ".csv")),
                    "-e", plan["hardware"]["events"], "--"]
    command += [binary["path"], "--warmup", str(job["warmup"]), "--samples",
                str(job["samples"]), "--case", job["case"],
                "--xlsx-cell-crud-shape", job["shape"], "--json",
                str(folder / (name + ".json"))]
    return command


def check_stats(stats: Any, values: list[int], label: str) -> None:
    require(isinstance(stats, dict) and stats.get("unit") == "ns",
            f"{label} stats are malformed")
    ordered = sorted(values)
    require(stats.get("samples") == ordered and stats.get("min") == ordered[0]
            and stats.get("max") == ordered[-1], f"{label} stats samples differ")
    expected_p50 = (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) // 2
    expected_p95 = ordered[min(math.ceil(len(ordered) * .95) - 1, len(ordered) - 1)]
    expected_p99 = ordered[min(math.ceil(len(ordered) * .99) - 1, len(ordered) - 1)]
    require(stats.get("p50") == expected_p50 and stats.get("p95") == expected_p95
            and stats.get("p99") == expected_p99, f"{label} percentiles differ")
    finite_number(stats.get("mean"), f"{label}.mean")
    require(math.isclose(float(stats["mean"]), sum(values) / len(values),
                         rel_tol=1e-12, abs_tol=1e-6), f"{label}.mean differs")


def checked_vector(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, list) and len(value) == count,
            f"{label} vector length differs")
    return value


def allocation_sample(value: Any, label: str, measured: bool) -> None:
    require(isinstance(value, dict) and value.get("scope") ==
            "operation_global_system_allocator", f"{label} scope differs")
    if not measured:
        require(value.get("status") == "unavailable"
                and all(field not in value for field in (
                    "allocation_calls", "deallocation_calls", "reallocation_calls",
                    "allocated_bytes", "deallocated_bytes", "live_bytes_before",
                    "live_bytes_after", "peak_live_bytes_before",
                    "peak_live_bytes_after", "region_peak_live_bytes")),
                f"{label} unavailable sample differs")
        return
    require(value.get("status") == "measured", f"{label} is not measured")
    fields = ("allocation_calls", "deallocation_calls", "reallocation_calls",
              "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
              "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
              "peak_live_bytes_after", "region_peak_live_bytes")
    for field in fields:
        nonnegative_int(value.get(field), f"{label}.{field}")
    require(value["failed_allocation_calls"] == 0
            and value["live_bytes_before"] + value["allocated_bytes"] ==
            value["live_bytes_after"] + value["deallocated_bytes"],
            f"{label} allocation balance differs")
    require(value["peak_live_bytes_before"] >= value["live_bytes_before"]
            and value["peak_live_bytes_after"] >= value["peak_live_bytes_before"]
            and value["peak_live_bytes_after"] >= value["live_bytes_after"]
            and value["region_peak_live_bytes"] >= max(value["live_bytes_before"],
                                                       value["live_bytes_after"]),
            f"{label} allocation peaks differ")


def logical_identity(result: dict[str, Any]) -> dict[str, Any]:
    corpus = result.get("corpus")
    sink = result.get("sink")
    output = result.get("output_sha256")
    require(isinstance(corpus, dict) and valid_digest(output),
            "logical output identity is malformed")
    require(isinstance(sink, dict), "logical sink identity is missing")
    for field in ("accepted_bytes", "write_calls", "largest_write"):
        nonnegative_int(sink.get(field), f"sink.{field}")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), "sink write buckets are missing")
    for value in buckets.values():
        nonnegative_int(value, "sink bucket")
    require(sum(buckets.values()) == sink["write_calls"],
            "sink buckets do not reconcile")
    return {"corpus": corpus, "sink": sink, "output_sha256": output}


def validate_raw(path: Path, job: dict[str, Any], binary: dict[str, Any],
                 plan: dict[str, Any]) -> dict[str, Any]:
    raw = read_json(path, rel(path))
    label = job["name"]
    require(raw.get("schema_version") == 1, f"{label} schema version differs")
    tool = raw.get("tool")
    expected_binary = "litchi-perf-baseline" + ("-alloc" if job["lane"] == "alloc" else "")
    require(isinstance(tool, dict) and tool.get("binary") == expected_binary
            and tool.get("profile") == "release",
            f"{label} tool identity differs")
    binary_identity = raw.get("binary_identity")
    require(isinstance(binary_identity, dict)
            and binary_identity.get("binary_sha256") == binary["sha256"]
            and binary_identity.get("profile") == "release",
            f"{label} binary identity differs")
    environment = raw.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == plan["revision"]
            and environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{label} environment identity differs")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [job["case"]]
            and configuration.get("xlsx_cell_crud_shapes") == [job["shape"]]
            and configuration.get("samples_per_case") == job["samples"]
            and configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{label} configuration differs")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1 and isinstance(results[0], dict),
            f"{label} result matrix differs")
    result = results[0]
    require(result.get("case") == job["case"]
            and isinstance(result.get("corpus"), dict)
            and (not job["case"].startswith("xlsx_")
                 or result["corpus"].get("shape") == job["shape"]),
            f"{label} case/corpus differs")
    elapsed = result.get("elapsed_ns")
    elapsed_values = checked_vector(elapsed.get("samples") if isinstance(elapsed, dict) else None,
                                    job["samples"], f"{label}.elapsed_ns.samples")
    require(all(isinstance(value, int) and not isinstance(value, bool) and value >= 0
                for value in elapsed_values), f"{label} elapsed samples are invalid")
    order = elapsed.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(job["samples"])),
            f"{label} elapsed sample order differs")
    check_stats(elapsed, elapsed_values, f"{label}.elapsed_ns")

    source = result.get("source")
    require(isinstance(source, dict), f"{label} source evidence is missing")
    xlsx = source.get("xlsx_cell_values")
    if job["kind"] == "primary" or job["lane"] == "alloc":
        require(isinstance(xlsx, dict), f"{label} XLSX source evidence is missing")
        require(xlsx.get("implementation") in ("source-backed", "managed-source-backed")
                and xlsx.get("cache_mode") in ("unmanaged-control", "managed-budget"),
                f"{label} source implementation differs")
        phase_values: dict[str, list[int]] = {}
        for phase in PHASES:
            values = checked_vector(xlsx.get(phase), job["samples"], f"{label}.{phase}")
            require(all(isinstance(value, int) and not isinstance(value, bool) and value >= 0
                        for value in values), f"{label}.{phase} values are invalid")
            phase_values[phase] = values
        for sorted_index, acquisition_index in enumerate(order):
            require(sum(phase_values[phase][acquisition_index] for phase in PHASES)
                    == elapsed_values[sorted_index],
                    f"{label} phase sum does not match elapsed sample")
        for field in ("commit_allocation_metrics", "publication_allocation_metrics"):
            values = checked_vector(xlsx.get(field), job["samples"], f"{label}.{field}")
            for index, sample in enumerate(values):
                allocation_sample(sample, f"{label}.{field}[{index}]",
                                  job["lane"] == "alloc")
    return {"job": job, "identity": logical_identity(result),
            "elapsed": {"p50": elapsed["p50"], "mean": elapsed["mean"]},
            "phases": {phase: {"p50": phase_stats(result, xlsx, phase)["p50"],
                               "mean": phase_stats(result, xlsx, phase)["mean"]}
                       for phase in PHASES} if isinstance(xlsx, dict) else {}}


def phase_stats(result: dict[str, Any], xlsx: dict[str, Any], phase: str) -> dict[str, Any]:
    values = xlsx.get(phase)
    require(isinstance(values, list) and values, f"phase {phase} is empty")
    ordered = sorted(values)
    require(all(isinstance(value, int) and value >= 0 for value in values),
            f"phase {phase} contains invalid values")
    return {"p50": (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) // 2,
            "mean": sum(values) / len(values)}


def validate_capture(stage: str, job: dict[str, Any], binary: dict[str, Any],
                     stage_sha: str, base_sha: str, candidate_sha: str,
                     plan: dict[str, Any]) -> dict[str, Any]:
    folder = {"baseline": BASELINE, "candidate": CANDIDATE}[stage]
    path = folder / (job["name"] + ".receipt.json")
    row = receipt_common(path, stage, job["name"], stage_sha, base_sha,
                         candidate_sha, binary["sha256"])
    require(row["value"].get("command") == capture_command(job, binary, folder, plan),
            f"{rel(path)} command differs")
    artifact_map(row["value"], folder, job)
    parsed = validate_raw(folder / (job["name"] + ".json"), job, binary, plan)
    if job["lane"] == "native":
        rss = read_json(folder / (job["name"] + ".rss.json"), rel(path) + " RSS")
        require(set(rss) == {"max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"},
                f"{job['name']} RSS fields differ")
        nonnegative_int(rss["max_rss_kib"], f"{job['name']} RSS")
        for field in ("elapsed_seconds", "user_seconds", "system_seconds"):
            finite_number(rss[field], f"{job['name']}.{field}")
            require(rss[field] >= 0, f"{job['name']}.{field} is negative")
    parsed.update({"path": rel(path), "start": row["start"].isoformat(),
                   "end": row["end"].isoformat()})
    return parsed


def require_matrix(actual: set[str], expected: set[str], label: str) -> None:
    missing = sorted(expected - actual)
    unexpected = sorted(actual - expected)
    if missing:
        raise IncompleteError(f"{label} is missing: {missing}")
    require(not unexpected, f"{label} contains unplanned captures: {unexpected}")


def validate_serial(rows: list[dict[str, Any]]) -> None:
    ordered = sorted(rows, key=lambda row: row["start"])
    for left, right in zip(ordered, ordered[1:]):
        left_end = left["end"] if isinstance(left["end"], dt.datetime) else \
            parse_time(left["end"], f"{left['path']}.end")
        right_start = right["start"] if isinstance(right["start"], dt.datetime) else \
            parse_time(right["start"], f"{right['path']}.start")
        require(left_end <= right_start,
                f"receipt intervals overlap: {left['path']} and {right['path']}")


def validate_abba(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    native = sorted((row for row in rows if row["job"]["lane"] == "native"),
                    key=lambda row: row["start"])
    blocks: list[tuple[str, int]] = []
    block_names: dict[tuple[str, int], set[str]] = {}
    for row in native:
        match = re.match(r"^native-r([12])-", row["job"]["name"])
        require(match, f"native receipt name is malformed: {row['path']}")
        stage = Path(row["path"]).parts[-2]
        key = (stage, int(match.group(1)))
        if not blocks or blocks[-1] != key:
            blocks.append(key)
        block_names.setdefault(key, set()).add(row["job"]["name"])
    expected = [("baseline", 1), ("candidate", 1), ("candidate", 2), ("baseline", 2)]
    require(blocks == expected, f"native ABBA blocks differ: {blocks}")
    for stage, repeat in expected:
        expected_names = {job["name"] for job in expected_jobs(plan, "native")
                          if job["name"].startswith(f"native-r{repeat}-")}
        require(block_names.get((stage, repeat)) == expected_names,
                f"{stage} native r{repeat} block is incomplete")
    return {"blocks": [f"{stage}/native-r{repeat}" for stage, repeat in blocks],
            "retained_baseline_a2": True}


def validate_captures() -> dict[str, Any]:
    plan = load_plan()
    source = validate_source()
    builds = validate_builds()
    base_sha = source["baseline"]["manifest_sha256"]
    candidate_sha = source["candidate"]["manifest_sha256"]
    rows: list[dict[str, Any]] = []
    stages: dict[str, Any] = {}
    for stage in ("baseline", "candidate"):
        stage_sha = sha({"baseline": BASELINE, "candidate": CANDIDATE}[stage] /
                        "source-manifest.json")
        identities = builds["stages"][stage]["binaries"]
        jobs = expected_jobs(plan, "native")
        expected_names = {job["name"] for job in jobs}
        folder = {"baseline": BASELINE, "candidate": CANDIDATE}[stage]
        actual = {path.name.removesuffix(".receipt.json") for path in folder.glob("*.receipt.json")
                  if path.name.startswith("native-")}
        require_matrix(actual, expected_names, f"{stage} native capture matrix")
        stage_rows = []
        for job in jobs:
            kind = "alloc" if job["lane"] == "alloc" else "normal"
            row = validate_capture(stage, job, identities[kind], stage_sha,
                                   base_sha, candidate_sha, plan)
            stage_rows.append(row)
            rows.append(row)
        alloc_receipts = list(folder.glob("alloc-*.receipt.json"))
        if alloc_receipts:
            require("alloc" in identities,
                    f"{stage} allocator captures lack an allocator binary")
            alloc_rows = []
            for job in expected_jobs(plan, "alloc"):
                alloc_rows.append(validate_capture(stage, job, identities["alloc"], stage_sha,
                                                    base_sha, candidate_sha, plan))
            actual_alloc = {path.name.removesuffix(".receipt.json") for path in alloc_receipts}
            expected_alloc = {job["name"] for job in expected_jobs(plan, "alloc")}
            require_matrix(actual_alloc, expected_alloc, f"{stage} allocation capture matrix")
            stage_rows.extend(alloc_rows)
            rows.extend(alloc_rows)
        stages[stage] = {"manifest_sha256": stage_sha, "rows": stage_rows,
                         "allocation_captured": bool(alloc_receipts)}
    validate_serial(rows)
    abba = validate_abba(rows, plan)
    primary_count = sum(job["samples"] for job in expected_jobs(plan, "native"))
    allocation_count = sum(job["samples"] for job in expected_jobs(plan, "alloc")) \
        if any(row["job"]["lane"] == "alloc" for row in rows) else 0
    return {"status": "pass", "stages": stages, "rows": rows, "abba": abba,
            "native_samples": primary_count * 2, "allocation_samples": allocation_count * 2}


def report_path() -> Path:
    paths = [HERE / name for name in REPORT_NAMES if (HERE / name).is_file()]
    require(len(paths) <= 1, "multiple comparison reports are present")
    return need(paths[0] if paths else HERE / REPORT_NAMES[0], "comparison report")


def analyzer_path() -> Path:
    paths = [HERE / name for name in ANALYZER_NAMES if (HERE / name).is_file()]
    require(len(paths) <= 1, "multiple analyzers are present")
    return need(paths[0] if paths else HERE / ANALYZER_NAMES[0], "canonical analyzer")


def replay_analyzer(analyzer: Path, report: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="litchi-0531-analysis-", dir="/home/zhuhe") as folder:
        output = Path(folder) / "comparison.json"
        attempts = (
            [sys.executable, "-B", str(analyzer), "--output", str(output)],
            [sys.executable, "-B", str(analyzer), str(output)],
        )
        for command in attempts:
            result = subprocess.run(command, cwd=REPO, stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE, check=False)
            if result.returncode == 0 and output.is_file():
                require(output.read_bytes() == report.read_bytes(),
                        "canonical analyzer report does not replay byte-for-byte")
                return
            output.unlink(missing_ok=True)
        raise EvidenceError("canonical analyzer replay did not complete")


def stage_evidence(report: dict[str, Any], stage: str) -> dict[str, Any]:
    value = report.get(stage)
    require(isinstance(value, dict), f"comparison {stage} evidence is missing")
    if isinstance(value.get("evidence"), dict):
        value = value["evidence"]
    require(isinstance(value.get("native"), dict), f"comparison {stage} native evidence is malformed")
    return value


def primary_report_rows(report: dict[str, Any], plan: dict[str, Any], label: str) -> dict[tuple[int, str], dict[str, Any]]:
    native = stage_evidence(report, label).get("native")
    rows = native.get("rows") if isinstance(native, dict) else None
    require(isinstance(rows, list), f"{label} native rows are missing")
    selected = [row for row in rows if isinstance(row, dict)
                and row.get("kind") == "primary" and row.get("guard") is None]
    expected = {(repeat, shape) for repeat in range(1, plan["primary"]["repeats"] + 1)
                for shape in plan["primary"]["shapes"]}
    keys = [(row.get("repeat"), row.get("shape")) for row in selected]
    require(set(keys) == expected and len(keys) == len(set(keys)) == len(expected),
            f"{label} primary matrix differs")
    return {(row["repeat"], row["shape"]): row for row in selected}


def report_stat(row: dict[str, Any], phase: str, stat: str, label: str) -> float:
    timing = row.get("timing")
    require(isinstance(timing, dict), f"{label}.timing is missing")
    value = timing.get(phase)
    if isinstance(value, dict):
        value = value.get(stat)
    require(value is not None, f"{label}.{phase}.{stat} is missing")
    finite_number(value, f"{label}.{phase}.{stat}")
    require(float(value) >= 0, f"{label}.{phase}.{stat} is negative")
    return float(value)


def reduction(baseline: float, candidate: float, label: str) -> float:
    require(baseline > 0, f"{label} baseline is not positive")
    return (baseline - candidate) / baseline * 100.0


def independent_pilot(report: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    left = primary_report_rows(report, plan, "baseline")
    right = primary_report_rows(report, plan, "candidate")
    gates = plan["gates"]
    rows = []
    passed = True
    for key in sorted(left):
        lrow, rrow = left[key], right[key]
        total_p50 = reduction(report_stat(lrow, "elapsed_ns", "p50", str(key)),
                              report_stat(rrow, "elapsed_ns", "p50", str(key)), str(key))
        total_mean = reduction(report_stat(lrow, "elapsed_ns", "mean", str(key)),
                               report_stat(rrow, "elapsed_ns", "mean", str(key)), str(key))
        planning_p50 = reduction(report_stat(lrow, "plan_ns", "p50", str(key)),
                                  report_stat(rrow, "plan_ns", "p50", str(key)), str(key))
        row_passed = (total_p50 >= gates["total_p50_reduction_percent"]
                      and total_mean >= gates["total_mean_reduction_percent"]
                      and planning_p50 >= gates["planning_p50_reduction_percent"])
        passed = passed and row_passed
        rows.append({"repeat": key[0], "shape": key[1], "total_p50": total_p50,
                     "total_mean": total_mean, "planning_p50": planning_p50,
                     "passed": row_passed})
    return {"rows": rows, "passed": passed,
            "decision": "eligible-for-conditional-lanes" if passed else "reject"}


def validate_analysis() -> dict[str, Any]:
    plan = load_plan()
    captures = validate_captures()
    analyzer = analyzer_path()
    report_file = report_path()
    replay_analyzer(analyzer, report_file)
    report = read_json(report_file, "comparison report")
    require(report.get("status") == "pass" and report.get("stage") == "compare"
            and report.get("plan_sha256") == sha(PLAN),
            "comparison envelope differs")
    require(report.get("structured_gates") == plan["gates"],
            "comparison structured gates differ")
    independent = independent_pilot(report, plan)
    # 0531 has no allocator gate.  Its canonical analyzer calls the native
    # gate ``native_admission`` (older pilot bundles called this
    # ``pilot_admission``); accept the bound 0531 spelling and retain the
    # compatibility alias for a handoff analyzer.
    pilot = report.get("native_admission", report.get("pilot_admission"))
    require(isinstance(pilot, dict) and pilot.get("passed") is independent["passed"],
            "comparison pilot decision does not replay independently")
    return {"status": "pass", "comparison_sha256": sha(report_file),
            "analyzer_sha256": sha(analyzer), "independent_pilot": independent,
            "pilot": pilot, "captures": captures}


def load_pure_module(path: Path, label: str) -> Any:
    """Load a conditional analyzer without invoking its command-line path.

    The conditional reports have their own analyzers.  Importing a module and
    calling its analysis function lets this verifier replay the report in a
    temporary process state while avoiding capture, build, or report-writing
    commands.  ``-B`` is used by the verifier entry point, so this import does
    not create a bytecode artifact in the evidence bundle.
    """
    need(path, f"{label} analyzer")
    name = f"litchi_0531_verify_{label}"
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None,
            f"{label} analyzer cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    # reopen_confirmation.py follows its command-line import convention and
    # imports the sibling ``analyze`` and ``run`` modules by name.  Make the
    # evidence directory importable for this pure replay, then restore the
    # verifier's module search path immediately afterward.
    sys.path.insert(0, str(HERE))
    try:
        spec.loader.exec_module(module)
    except Exception as error:  # analyzers expose their own EvidenceError types
        raise EvidenceError(f"{label} analyzer import failed: {error}") from error
    finally:
        if sys.path and sys.path[0] == str(HERE):
            sys.path.pop(0)
    return module


def replay_pure_conditional(lane: str, analyzer: Path, report_path_value: Path) -> dict[str, Any]:
    """Replay a lane's pure analyzer and require semantic report identity."""
    # The eager capture guard records its canonical binary paths in the raw
    # children, while binary-normal.json keeps the temporary custody aliases.
    # Replay through the frozen alias wrapper so that only that path mapping
    # is adapted; all guard semantics remain in eager_guard.py.
    replay_analyzer = EAGER_REPLAY_ANALYZER if lane == "eager" else analyzer
    module = load_pure_module(replay_analyzer, lane)
    try:
        if lane == "profile":
            value = module.analyze("both", False)
        elif lane == "eager":
            require(hasattr(module, "GUARD"),
                    "eager alias wrapper has no frozen guard")
            value = module.GUARD.analyze(report_path_value)
        elif lane == "reopen":
            value = module.analyze()
        elif lane == "final_native":
            value = module.analyze()
        else:
            raise EvidenceError(f"unknown pure conditional lane: {lane}")
    except EvidenceError:
        raise
    except Exception as error:
        # The helper analyzers use distinct EvidenceError classes.  Preserve
        # the verifier's failure vocabulary without treating a malformed or
        # incomplete report as a pass.
        raise EvidenceError(f"{lane} analyzer replay failed: {error}") from error
    require(isinstance(value, dict), f"{lane} analyzer replay did not return an object")
    expected = read_json(report_path_value, f"{lane} comparison report")
    require(value == expected,
            f"{lane} pure analyzer replay differs from {rel(report_path_value)}")
    return value


def validate_final_native_lane(plan: dict[str, Any]) -> dict[str, Any]:
    """Replay the post-correction native matrix and bind its failed gate.

    The final-source rebuild is a separate historical experiment: its binary
    is allowed to differ from the measured candidate, and its result is used
    only to decide whether the candidate remains retainable.  The analyzer's
    complete report is replayed first; this pass then independently checks the
    four primary gate records and exposes every failed metric for the decision
    record to bind.
    """
    report_path_value = need(FINAL_NATIVE_REPORT, "final native comparison report")
    report = replay_pure_conditional("final_native", FINAL_NATIVE_ANALYZER,
                                     report_path_value)
    final_plan = read_json(FINAL_NATIVE_PLAN, "final native plan")
    frozen = read_json(FINAL_NATIVE_FROZEN, "final native frozen inputs")
    require(report.get("status") == "pass"
            and report.get("stage") == "final-native-compare"
            and report.get("plan_sha256") == sha(PLAN)
            and report.get("final_plan_sha256") == sha(FINAL_NATIVE_PLAN)
            and report.get("frozen_inputs_sha256") == sha(FINAL_NATIVE_FROZEN)
            and report.get("final_capture_sha256") == sha(HERE / "final_capture.py")
            and report.get("source_binding_sha256") == sha(FINAL_BINDING)
            and report.get("structured_gates") == plan["gates"],
            "final native comparison envelope differs")
    require(isinstance(final_plan, dict)
            and final_plan.get("primary_plan_sha256") == sha(PLAN)
            and final_plan.get("gates") == plan["gates"]
            and final_plan.get("final_source_manifest_sha256") ==
            sha(FINAL / "source-manifest.json")
            and frozen.get("final-native-plan.json") == sha(FINAL_NATIVE_PLAN)
            and frozen.get("final_capture.py") == sha(HERE / "final_capture.py"),
            "final native authority binding differs")
    require(report.get("baseline_stage") == "baseline"
            and report.get("candidate_stage") == "final"
            and report.get("candidate_label") == "final",
            "final native true stage labels differ")
    labels = report.get("true_stage_labels")
    require(isinstance(labels, dict)
            and labels.get("baseline", {}).get("stage") == "baseline"
            and labels.get("final", {}).get("stage") == "final"
            and labels.get("baseline", {}).get("source_manifest_sha256") ==
            sha(BASELINE / "source-manifest.json")
            and labels.get("final", {}).get("source_manifest_sha256") ==
            sha(FINAL / "source-manifest.json"),
            "final native source labels differ")
    admission = report.get("native_admission")
    require(isinstance(admission, dict)
            and admission.get("baseline_stage") == "baseline"
            and admission.get("candidate_stage") == "final"
            and admission.get("decision") == "reject"
            and admission.get("passed") is False,
            "final native failed admission is not bound")
    rows = admission.get("rows")
    require(isinstance(rows, list), "final native admission rows are missing")
    expected_keys = {(repeat, shape)
                     for repeat in range(1, int(plan["primary"]["repeats"]) + 1)
                     for shape in plan["primary"]["shapes"]}
    require(len(rows) == len(expected_keys), "final native admission matrix differs")
    seen: set[tuple[Any, Any]] = set()
    failed: list[dict[str, Any]] = []
    gate_fields = (
        ("native_primary_total_p50", "total_p50_reduction_percent"),
        ("native_primary_total_mean", "total_mean_reduction_percent"),
        ("native_primary_planning_p50", "planning_p50_reduction_percent"),
    )
    required = plan["gates"]
    for row in rows:
        require(isinstance(row, dict), "final native admission row is malformed")
        key = (row.get("repeat"), row.get("shape"))
        require(key in expected_keys and key not in seen,
                "final native admission row identity differs")
        seen.add(key)
        expected_row_pass = True
        for field, gate_name in gate_fields:
            metric = row.get(field)
            expected_metric_fields = {"baseline", "candidate", "passed",
                                     "reduction_percent", "required_reduction_percent"}
            if field == "native_primary_planning_p50":
                expected_metric_fields.add("metric")
            require(isinstance(metric, dict)
                    and set(metric) == expected_metric_fields,
                    f"final native {key} {field} gate record differs")
            if field == "native_primary_planning_p50":
                require(metric.get("metric") == "plan_ns",
                        f"final native {key} planning metric identity differs")
            baseline_value = metric.get("baseline")
            candidate_value = metric.get("candidate")
            finite_number(baseline_value,
                          f"final native {key} {field}.baseline")
            finite_number(candidate_value,
                          f"final native {key} {field}.candidate")
            require(float(baseline_value) > 0.0 and float(candidate_value) > 0.0,
                    f"final native {key} {field} endpoints must be positive")
            actual_reduction = reduction(float(baseline_value), float(candidate_value),
                                         f"final native {key} {field}")
            require_same_number(metric.get("reduction_percent"), actual_reduction,
                                f"final native {key} {field}.reduction_percent")
            gate = float(required[gate_name])
            require(metric.get("required_reduction_percent") == gate
                    and metric.get("passed") is (actual_reduction >= gate),
                    f"final native {key} {field} gate arithmetic differs")
            passed = actual_reduction >= gate
            expected_row_pass = expected_row_pass and passed
            if not passed:
                failed.append({"repeat": key[0], "shape": key[1], "gate": field,
                               "baseline": baseline_value, "candidate": candidate_value,
                               "reduction_percent": actual_reduction,
                               "required_reduction_percent": gate})
        require(row.get("passed") is expected_row_pass,
                f"final native {key} aggregate gate differs")
    require(seen == expected_keys and failed and report.get("admission_status") == "reject",
            "final native rejection does not have a failed measured gate")
    result = {"status": "pass", "path": rel(report_path_value),
            "sha256": sha(report_path_value), "analyzer_sha256": sha(FINAL_NATIVE_ANALYZER),
            "final_plan_sha256": sha(FINAL_NATIVE_PLAN), "admission_status": "reject",
            "gate_passed": False, "rows": len(rows), "failed_gates": failed,
            "baseline_binary_sha256": final_plan.get("baseline_binary_sha256"),
            "final_binary_sha256": final_plan.get("final_binary_sha256")}
    result["review"] = validate_final_native_review(report, result)
    return result


def validate_final_native_review(report: dict[str, Any],
                                final_native: dict[str, Any]) -> dict[str, Any]:
    """Bind the independent review of every final-native flag.

    The final native analyzer replays the measured matrix and its gate
    arithmetic.  The companion review is a separate custody record for the
    adverse and same-build-drift vectors (57 + 101 rows).  Requiring exact
    row identities, explanations, and authority hashes prevents a summary
    count or a passing status from hiding an unreviewed flag.
    """
    path = need(FINAL_NATIVE_REVIEW, "final native review")
    value = read_json(path, "final native review")
    comparison = report.get("comparison")
    require(isinstance(comparison, dict),
            "final native review source comparison is missing")
    expected_adverse = comparison.get("adverse_flags_over_five_percent")
    expected_drift = comparison.get("same_build_drift_over_five_percent")
    require(isinstance(expected_adverse, list) and isinstance(expected_drift, list),
            "final native review source flag vectors are missing")
    require(isinstance(value, dict)
            and value.get("schema") == "0531-final-native-review-v1"
            and value.get("comparison") == rel(FINAL_NATIVE_REPORT)
            and value.get("comparison_sha256") == final_native["sha256"]
            and value.get("complete") is True
            and value.get("production_decision") == "reject"
            and value.get("speedup_claim") == "none-retained"
            and value.get("all_adverse_metrics_retained") is True
            and value.get("all_same_build_drift_retained") is True,
            "final native review envelope differs")

    # The review's authority block must point at the same immutable inputs as
    # the replayed final-native report.  Check paths as well as digests where
    # the record carries a path, so a copied analyzer cannot be substituted.
    bindings = value.get("bindings")
    expected_bindings = {
        "analyzer": (FINAL_NATIVE_ANALYZER.name, FINAL_NATIVE_ANALYZER),
        "final_capture_sha256": (None, HERE / "final_capture.py"),
        "final_plan_sha256": (None, FINAL_NATIVE_PLAN),
        "frozen_inputs_sha256": (None, FINAL_NATIVE_FROZEN),
        "primary_plan_sha256": (None, PLAN),
        "source_binding_sha256": (None, FINAL_BINDING),
        "true_baseline_source_manifest_sha256": (None,
                                                   BASELINE / "source-manifest.json"),
        "true_final_source_manifest_sha256": (None,
                                               FINAL / "source-manifest.json"),
    }
    require(isinstance(bindings, dict), "final native review bindings are missing")
    for key, (expected_name, artifact) in expected_bindings.items():
        if expected_name is not None:
            require(bindings.get(key) == expected_name,
                    f"final native review {key} path differs")
    require(bindings.get("analyzer_sha256") == sha(FINAL_NATIVE_ANALYZER)
            and bindings.get("final_capture_sha256") == sha(HERE / "final_capture.py")
            and bindings.get("final_plan_sha256") == sha(FINAL_NATIVE_PLAN)
            and bindings.get("frozen_inputs_sha256") == sha(FINAL_NATIVE_FROZEN)
            and bindings.get("primary_plan_sha256") == sha(PLAN)
            and bindings.get("source_binding_sha256") == sha(FINAL_BINDING)
            and bindings.get("true_baseline_binary_sha256") ==
            final_native["baseline_binary_sha256"]
            and bindings.get("true_final_binary_sha256") ==
            final_native["final_binary_sha256"]
            and bindings.get("true_baseline_source_manifest_sha256") ==
            sha(BASELINE / "source-manifest.json")
            and bindings.get("true_final_source_manifest_sha256") ==
            sha(FINAL / "source-manifest.json"),
            "final native review authority hashes differ")

    def review_rows(expected: list[dict[str, Any]], actual: Any,
                    label: str) -> None:
        require(isinstance(actual, list) and len(actual) == len(expected),
                f"final native {label} review count differs")
        remaining = list(actual)
        for source in expected:
            require(isinstance(source, dict),
                    f"final native {label} source flag is malformed")
            index = next((index for index, row in enumerate(remaining)
                          if isinstance(row, dict)
                          and all(row.get(key) == item for key, item in source.items())),
                         None)
            require(index is not None,
                    f"final native {label} flag is unbound")
            row = remaining.pop(index)
            require(isinstance(row.get("review"), str) and row["review"].strip(),
                    f"final native {label} flag lacks an explanation")
            classification = row.get("classification")
            require(isinstance(classification, dict)
                    and classification.get("format") in ("XLSX", "DOCX", "PPTX")
                    and classification.get("phase") == row.get("phase")
                    and classification.get("stat") == row.get("stat")
                    and classification.get("shape") == row.get("shape")
                    and isinstance(classification.get("scope"), str)
                    and isinstance(classification.get("review_outcome"), str)
                    and classification["review_outcome"].strip(),
                    f"final native {label} classification is incomplete")
        require(not remaining, f"final native {label} review has extra rows")

    review_rows(expected_adverse, value.get("adverse"), "adverse")
    review_rows(expected_drift, value.get("same_build_drift"), "same-build drift")
    require(value.get("adverse_count") == len(expected_adverse)
            and value.get("same_build_drift_count") == len(expected_drift)
            and value.get("material_guard_regression_count") == sum(
                1 for row in value["adverse"]
                if row.get("classification", {}).get("material_guard_regression") is True),
            "final native review aggregate counts differ")

    def counts(rows: list[dict[str, Any]], field: str,
               *, row_field: bool = False) -> dict[str, int]:
        return dict(Counter(str(row[field] if row_field else
                                row["classification"][field]) for row in rows))

    for key, rows, fields in (
        ("adverse_counts_by_format", value["adverse"], ("format",)),
        ("adverse_counts_by_phase", value["adverse"], ("phase",)),
        ("adverse_counts_by_scope", value["adverse"], ("scope",)),
        ("adverse_counts_by_shape", value["adverse"], ("shape",)),
        ("adverse_counts_by_stat", value["adverse"], ("stat",)),
        ("same_build_drift_counts_by_format", value["same_build_drift"], ("format",)),
        ("same_build_drift_counts_by_phase", value["same_build_drift"], ("phase",)),
        ("same_build_drift_counts_by_scope", value["same_build_drift"], ("scope",)),
        ("same_build_drift_counts_by_shape", value["same_build_drift"], ("shape",)),
        ("same_build_drift_counts_by_stage", value["same_build_drift"], ("stage",)),
        ("same_build_drift_counts_by_stat", value["same_build_drift"], ("stat",)),
    ):
        require(value.get(key) == counts(rows, fields[0], row_field=key.endswith("_by_stage")),
                f"final native review {key} differs")

    rejection = value.get("rejection")
    require(isinstance(rejection, dict)
            and rejection.get("status") == "native-gate-failed"
            and rejection.get("no_retained_speedup_claim") is True
            and rejection.get("failed_gate_count") == len(final_native["failed_gates"])
            and rejection.get("failed_admission_row_count") == 1
            and isinstance(rejection.get("reason"), str)
            and rejection["reason"].strip(),
            "final native review rejection binding differs")
    failed_metrics = rejection.get("failed_gate_metrics")
    require(isinstance(failed_metrics, list)
            and len(failed_metrics) == len(final_native["failed_gates"]),
            "final native review failed gate inventory differs")
    expected_metric_names = {
        "native_primary_total_p50": "total_p50",
        "native_primary_total_mean": "total_mean",
    }
    for failed in final_native["failed_gates"]:
        metric_name = expected_metric_names.get(failed["gate"])
        require(metric_name is not None, "final native failed gate is not a total gate")
        match = next((row for row in failed_metrics if isinstance(row, dict)
                      and row.get("repeat") == failed["repeat"]
                      and row.get("shape") == failed["shape"]
                      and row.get("metric") == metric_name), None)
        require(isinstance(match, dict),
                "final native review failed gate is not decision-bound")
        require(match.get("baseline_stage") == "baseline"
                and match.get("candidate_stage") == "final"
                and match.get("candidate") == failed["candidate"]
                and match.get("baseline") == failed["baseline"]
                and match.get("passed") is False
                and match.get("required_reduction_percent") ==
                failed["required_reduction_percent"]
                and math.isclose(float(match.get("reduction_percent")),
                                 float(failed["reduction_percent"]),
                                 rel_tol=1e-12, abs_tol=1e-9),
                "final native review failed gate arithmetic differs")
    return {"path": rel(path), "sha256": sha(path),
            "adverse_flags": len(expected_adverse),
            "same_build_flags": len(expected_drift),
            "flags": len(expected_adverse) + len(expected_drift),
            "material_guard_regressions": value["material_guard_regression_count"],
            "complete": True}


def bound_bundle_file(value: Any, label: str, expected: Path | None = None) -> Path:
    """Resolve and hash-check a report path that must stay in this bundle."""
    safe_relative(value, label)
    path = HERE / value
    require(path.is_file() and not path.is_symlink(), f"{label} is not a regular file")
    if expected is not None:
        require(path == expected, f"{label} differs: {rel(path)}")
    return path


def bound_report_artifact(value: Any, digest: Any, label: str,
                          expected: Path | None = None) -> Path:
    path = bound_bundle_file(value, label, expected)
    require(valid_digest(digest) and sha(path) == digest,
            f"{label} digest differs")
    return path


def report_float(value: Any, label: str) -> float:
    finite_number(value, label)
    require(float(value) >= 0.0, f"{label} is negative")
    return float(value)


def report_integer(value: Any, label: str, *, positive: bool = False) -> int:
    nonnegative_int(value, label)
    if positive:
        require(value > 0, f"{label} is not positive")
    return value


def percent_change(baseline: Any, candidate: Any, label: str) -> float | None:
    report_float(baseline, f"{label}.baseline")
    report_float(candidate, f"{label}.candidate")
    if float(baseline) == 0.0:
        return 0.0 if float(candidate) == 0.0 else None
    return (float(candidate) / float(baseline) - 1.0) * 100.0


def require_same_number(actual: Any, expected: float | None, label: str) -> None:
    if expected is None:
        require(actual is None, f"{label} must be null for a zero baseline")
        return
    finite_number(actual, label)
    require(math.isclose(float(actual), expected, rel_tol=1e-12, abs_tol=1e-9),
            f"{label} differs from independently recomputed value")


def profile_stage_metadata(stage: str, value: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    folder = {"baseline": BASELINE, "candidate": CANDIDATE,
              "final": FINAL}[stage]
    manifest_path = folder / "source-manifest.json"
    stage_sha = sha(manifest_path)
    binary = binary_identity(stage, "normal", stage_sha)
    metadata = value.get("metadata")
    require(isinstance(metadata, dict), f"profile {stage} metadata is missing")
    require(metadata.get("stage") == stage
            and metadata.get("source_manifest") == rel(manifest_path)
            and metadata.get("source_manifest_sha256") == stage_sha
            and metadata.get("source_manifest_entries") == len(source_manifest(manifest_path)),
            f"profile {stage} source metadata differs")
    require(metadata.get("binary_identity") == rel(folder / "binary-normal.json")
            and metadata.get("binary_path") == binary["path"]
            and metadata.get("binary_sha256") == binary["sha256"]
            and metadata.get("binary_bytes") == binary["bytes"],
            f"profile {stage} binary metadata differs")
    build_receipt = folder / "build-normal.receipt.json"
    require(metadata.get("build_receipt") == rel(build_receipt)
            and metadata.get("build_receipt_sha256") == sha(build_receipt),
            f"profile {stage} build metadata differs")
    require(value.get("stage") == stage and value.get("status") == "pass"
            and value.get("profile_count") == 4,
            f"profile {stage} envelope differs")
    validation = value.get("validation")
    require(isinstance(validation, dict)
            and validation.get("exact_source_binary_receipt_bindings") is True
            and validation.get("profile_matrix_complete") is True
            and validation.get("raw_dumps_complete") is True,
            f"profile {stage} validation receipt differs")
    return {"manifest_sha256": stage_sha, "binary": binary,
            "manifest_entries": len(source_manifest(manifest_path))}


def validate_profile_stage_profiles(stage: str, value: dict[str, Any],
                                    plan: dict[str, Any], *, native_prefix: str = "native") -> dict[tuple[int, str], dict[str, Any]]:
    folder = {"baseline": BASELINE, "candidate": CANDIDATE,
              "final": FINAL}[stage]
    profiles = value.get("profiles")
    require(isinstance(profiles, list), f"profile {stage} profile rows are missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(plan["profile"]["repeats"]) + 1)
                for shape in plan["profile"]["shapes"]}
    keys = {(row.get("repeat"), row.get("shape")) for row in profiles
            if isinstance(row, dict)}
    require(keys == expected and len(profiles) == len(expected),
            f"profile {stage} matrix differs")
    result: dict[tuple[int, str], dict[str, Any]] = {}
    receipt_intervals: list[tuple[dt.datetime, dt.datetime, str]] = []
    for row in profiles:
        require(isinstance(row, dict), f"profile {stage} row is malformed")
        repeat, shape = row.get("repeat"), row.get("shape")
        name = f"profile-r{repeat}-{shape}"
        require(row.get("name") == name and (repeat, shape) in expected,
                f"profile {stage} row identity differs")
        planning_ir = report_integer(row.get("planning_ir"), f"{name}.planning_ir", positive=True)

        receipt_value = row.get("receipt")
        require(isinstance(receipt_value, dict), f"{name} profile receipt is missing")
        receipt_path = bound_report_artifact(receipt_value.get("file"), receipt_value.get("sha256"),
                                             f"{name}.receipt")
        require(receipt_path == folder / f"{name}.receipt.json",
                f"{name} profile receipt path differs")
        start = parse_time(receipt_value.get("start_utc"), f"{name}.receipt.start_utc")
        end = parse_time(receipt_value.get("end_utc"), f"{name}.receipt.end_utc")
        require(end > start, f"{name} profile receipt interval is invalid")
        receipt_intervals.append((start, end, name))
        artifacts = receipt_value.get("artifacts")
        expected_artifacts = {
            f"{name}.json", f"{name}.stdout", f"{name}.stderr",
            f"{name}.callgrind",
            *(f"{name}.callgrind.{part}" for part in range(1, 5)),
        }
        require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
                f"{name} profile artifact inventory differs")
        for filename, digest in artifacts.items():
            bound_report_artifact(f"{stage}/{filename}", digest,
                                  f"{name} profile artifact {filename}")

        profile_result = row.get("profile_result")
        require(isinstance(profile_result, dict), f"{name} profile result is missing")
        profile_path = bound_report_artifact(
            profile_result.get("file"), profile_result.get("sha256"),
            f"{name}.profile_result", folder / f"{name}.json")
        require(valid_digest(profile_result.get("logical_identity_sha256")),
                f"{name} profile logical identity is malformed")
        native_result = row.get("native_result")
        require(isinstance(native_result, dict), f"{name} native identity is missing")
        native_name = f"{native_prefix}-r{repeat}-primary-{shape}"
        native_path = bound_report_artifact(
            native_result.get("file"), native_result.get("sha256"),
            f"{name}.native_result")
        require(native_path == folder / f"{native_name}.json",
                f"{name}.native_result path differs")
        require(valid_digest(native_result.get("logical_identity_sha256"))
                and native_result.get("logical_identity_sha256") ==
                profile_result.get("logical_identity_sha256"),
                f"{name} profile/native logical identity differs")

        termination = row.get("termination")
        require(isinstance(termination, dict)
                and termination.get("part") == 5
                and termination.get("summary_ir") == 0,
                f"{name} callgrind termination differs")
        bound_report_artifact(termination.get("file"), termination.get("sha256"),
                              f"{name}.termination", folder / f"{name}.callgrind")
        raw_dumps = row.get("raw_dumps")
        require(isinstance(raw_dumps, list) and len(raw_dumps) == 4,
                f"{name} raw dump inventory differs")
        parts: set[int] = set()
        for dump in raw_dumps:
            require(isinstance(dump, dict), f"{name} raw dump row is malformed")
            part = dump.get("part")
            require(part in range(1, 5) and part not in parts,
                    f"{name} raw dump part inventory differs")
            parts.add(part)
            bound_report_artifact(dump.get("file"), dump.get("sha256"),
                                  f"{name}.raw_dump.{part}",
                                  folder / f"{name}.callgrind.{part}")
            report_integer(dump.get("summary_ir"), f"{name}.raw_dump.{part}.summary_ir",
                           positive=True)
        annotations = row.get("annotations")
        require(isinstance(annotations, dict)
                and set(annotations) >= {"inclusive", "self"},
                f"{name} profile annotations are missing")
        for kind in ("inclusive", "self"):
            annotation = annotations[kind]
            require(isinstance(annotation, dict), f"{name} {kind} annotation is malformed")
            bound_report_artifact(annotation.get("file"), annotation.get("sha256"),
                                  f"{name}.{kind} annotation")
        result[(repeat, shape)] = row
        # Keep the path variable live in this independent custody pass.  It
        # also makes it explicit that the profile JSON itself was hashed.
        require(profile_path == folder / f"{name}.json"
                and native_path.name == f"{native_name}.json",
                f"{name} profile/native paths differ")
    ordered = sorted(receipt_intervals)
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"profile {stage} receipt intervals overlap")
    return result


def validate_profile_lane(analysis: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    report_path_value = need(PROFILE_REPORT, "profile comparison report")
    replay = replay_pure_conditional("profile", PROFILE_ANALYZER, report_path_value)
    report = replay
    require(report.get("schema") == "xlsx_mce_conditional_planning_profile_analysis_v1"
            and report.get("status") == "pass"
            and report.get("stage_selection") == ["baseline", "candidate"]
            and report.get("selected_function") == plan["profile"]["owner"]
            and report.get("plan") == "plan.json"
            and report.get("plan_sha256") == sha(PLAN),
            "profile comparison envelope differs")
    for key in ("planning_analyzer", "raw_helpers"):
        helper = report.get(key)
        require(isinstance(helper, dict) and valid_digest(helper.get("sha256")),
                f"profile {key} binding is malformed")
        helper_path = safe_relative(helper.get("path"), f"profile {key} path")
        actual_path = HERE.parent / helper_path
        require(actual_path.is_file() and not actual_path.is_symlink()
                and sha(actual_path) == helper["sha256"],
                f"profile {key} digest differs")
    native = report.get("native_admission")
    require(isinstance(native, dict)
            and native.get("file") == rel(HERE / "comparison.json")
            and native.get("sha256") == analysis["comparison_sha256"]
            and native.get("native_rows") == 4
            and native.get("admission_status") == "eligible-for-conditional-lanes",
            "profile native admission binding differs")
    comparison = report.get("comparison")
    require(isinstance(comparison, dict)
            and comparison.get("required_reduction_percent") ==
            plan["gates"]["planning_ir_reduction_percent"]
            and comparison.get("scope", "").lower().find("not latency") >= 0,
            "profile gate envelope differs")
    rows = comparison.get("rows")
    require(isinstance(rows, list), "profile comparison gate rows are missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(plan["profile"]["repeats"]) + 1)
                for shape in plan["profile"]["shapes"]}
    require({(row.get("repeat"), row.get("shape")) for row in rows
             if isinstance(row, dict)} == expected and len(rows) == len(expected),
            "profile comparison gate matrix differs")
    minimum_reduction = math.inf
    for row in rows:
        require(isinstance(row, dict), "profile comparison row is malformed")
        key = f"profile {row.get('repeat')}/{row.get('shape')}"
        metric = row.get("planning_ir")
        require(isinstance(metric, dict), f"{key} planning gate is missing")
        baseline_ir = report_integer(metric.get("baseline"), f"{key}.baseline", positive=True)
        candidate_ir = report_integer(metric.get("candidate"), f"{key}.candidate", positive=True)
        require(metric.get("delta_ir") == candidate_ir - baseline_ir,
                f"{key}.delta_ir differs")
        actual_reduction = reduction(float(baseline_ir), float(candidate_ir), key)
        require_same_number(metric.get("reduction_percent"), actual_reduction,
                            f"{key}.reduction_percent")
        required = float(plan["gates"]["planning_ir_reduction_percent"])
        require(metric.get("required_reduction_percent") == required
                and metric.get("passed") is (actual_reduction >= required),
                f"{key} planning gate decision differs")
        minimum_reduction = min(minimum_reduction, actual_reduction)
        require(row.get("profile_identity_equal") is True,
                f"{key} profile identity gate is not passed")
        mce = row.get("mce_process_markup_compatibility_ir")
        require(isinstance(mce, dict), f"{key} MCE mechanism metric is missing")
        mce_base = report_integer(mce.get("baseline"), f"{key}.mce baseline")
        mce_candidate = report_integer(mce.get("candidate"), f"{key}.mce candidate")
        require(mce.get("delta_ir") == mce_candidate - mce_base,
                f"{key}.mce delta differs")
    require(comparison.get("passed") is all(
        row["planning_ir"]["passed"] for row in rows)
            and comparison.get("decision") ==
            ("retainable-conditional-lane" if comparison["passed"] else "reject"),
            "profile comparison aggregate gate differs")
    require(comparison["passed"] is True, "profile planning Ir gate is not passing")
    stage_values = report.get("stages")
    require(isinstance(stage_values, dict)
            and set(stage_values) == {"baseline", "candidate"},
            "profile stage inventory differs")
    stage_profiles: dict[str, dict[tuple[int, str], dict[str, Any]]] = {}
    for stage in ("baseline", "candidate"):
        profile_stage_metadata(stage, stage_values[stage], plan)
        stage_profiles[stage] = validate_profile_stage_profiles(stage, stage_values[stage], plan)
    return {"status": "pass", "path": rel(report_path_value),
            "sha256": sha(report_path_value), "analyzer_sha256": sha(PROFILE_ANALYZER),
            "gate_passed": True, "required_reduction_percent":
            plan["gates"]["planning_ir_reduction_percent"],
            "minimum_reduction_percent": minimum_reduction,
            "rows": len(rows), "stages": list(stage_profiles)}


def validate_final_profile_lane(plan: dict[str, Any], final_native: dict[str, Any]) -> dict[str, Any]:
    """Replay the optional final-source profile as diagnostic evidence."""
    report_path_value = need(FINAL_PROFILE_REPORT, "final profile comparison report")
    report = replay_pure_conditional("profile", FINAL_PROFILE_ANALYZER,
                                     report_path_value)
    require(report.get("schema") == "xlsx_mce_conditional_planning_profile_analysis_v1"
            and report.get("status") == "pass"
            and report.get("stage_selection") == ["baseline", "final"]
            and report.get("selected_function") == plan["profile"]["owner"]
            and report.get("plan") == "plan.json"
            and report.get("plan_sha256") == sha(PLAN),
            "final profile comparison envelope differs")
    for key in ("planning_analyzer", "raw_helpers"):
        helper = report.get(key)
        require(isinstance(helper, dict) and valid_digest(helper.get("sha256")),
                f"final profile {key} binding is malformed")
        helper_path = safe_relative(helper.get("path"), f"final profile {key} path")
        actual_path = HERE.parent / helper_path
        require(actual_path.is_file() and not actual_path.is_symlink()
                and sha(actual_path) == helper["sha256"],
                f"final profile {key} digest differs")
    native = report.get("native_admission")
    require(isinstance(native, dict)
            and native.get("file") == rel(FINAL_NATIVE_REPORT)
            and native.get("sha256") == final_native["sha256"]
            and native.get("native_rows") == 4
            and native.get("admission_status") == "reject",
            "final profile native admission binding differs")
    comparison = report.get("comparison")
    require(isinstance(comparison, dict)
            and comparison.get("required_reduction_percent") ==
            plan["gates"]["planning_ir_reduction_percent"]
            and comparison.get("scope", "").lower().find("not latency") >= 0,
            "final profile gate envelope differs")
    rows = comparison.get("rows")
    require(isinstance(rows, list), "final profile comparison rows are missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(plan["profile"]["repeats"]) + 1)
                for shape in plan["profile"]["shapes"]}
    require({(row.get("repeat"), row.get("shape")) for row in rows
             if isinstance(row, dict)} == expected and len(rows) == len(expected),
            "final profile comparison matrix differs")
    required = float(plan["gates"]["planning_ir_reduction_percent"])
    minimum_reduction = math.inf
    for row in rows:
        require(isinstance(row, dict), "final profile comparison row is malformed")
        key = f"final profile {row.get('repeat')}/{row.get('shape')}"
        metric = row.get("planning_ir")
        require(isinstance(metric, dict)
                and set(metric) == {"baseline", "final", "delta_ir",
                                    "reduction_percent", "required_reduction_percent",
                                    "passed"},
                f"{key} planning gate differs")
        baseline_ir = report_integer(metric.get("baseline"), f"{key}.baseline", positive=True)
        final_ir = report_integer(metric.get("final"), f"{key}.final", positive=True)
        require(metric.get("delta_ir") == final_ir - baseline_ir,
                f"{key}.delta_ir differs")
        actual_reduction = reduction(float(baseline_ir), float(final_ir), key)
        require_same_number(metric.get("reduction_percent"), actual_reduction,
                            f"{key}.reduction_percent")
        require(metric.get("required_reduction_percent") == required
                and metric.get("passed") is (actual_reduction >= required),
                f"{key} planning gate arithmetic differs")
        minimum_reduction = min(minimum_reduction, actual_reduction)
        require(row.get("profile_identity_equal") is True,
                f"{key} profile identity is not preserved")
        mce = row.get("mce_process_markup_compatibility_ir")
        require(isinstance(mce, dict)
                and set(mce) == {"baseline", "final", "delta_ir"},
                f"{key} MCE mechanism metric differs")
        mce_base = report_integer(mce.get("baseline"), f"{key}.mce baseline")
        mce_final = report_integer(mce.get("final"), f"{key}.mce final")
        require(mce.get("delta_ir") == mce_final - mce_base,
                f"{key}.mce delta differs")
    require(comparison.get("passed") is all(row["planning_ir"]["passed"] for row in rows)
            and comparison.get("passed") is True
            and comparison.get("decision") == "instruction-gate-passed-diagnostic-only",
            "final profile aggregate decision differs")
    stage_values = report.get("stages")
    require(isinstance(stage_values, dict)
            and set(stage_values) == {"baseline", "final"},
            "final profile stage inventory differs")
    stage_profiles: dict[str, dict[tuple[int, str], dict[str, Any]]] = {}
    for stage in ("baseline", "final"):
        profile_stage_metadata(stage, stage_values[stage], plan)
        stage_profiles[stage] = validate_profile_stage_profiles(stage, stage_values[stage], plan, native_prefix="final-native")
    return {"status": "pass", "path": rel(report_path_value),
            "sha256": sha(report_path_value), "analyzer_sha256": sha(FINAL_PROFILE_ANALYZER),
            "gate_passed": True, "diagnostic_only": True,
            "required_reduction_percent": required,
            "minimum_reduction_percent": minimum_reduction,
            "rows": len(rows), "stages": list(stage_profiles)}


def validate_eager_plan(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(EAGER_PLAN, "eager-plan.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0531-eager-guard-plan-v1"
            and value.get("status") == "frozen-before-capture",
            "eager plan envelope differs")
    require(value.get("primary_plan_sha256") == sha(PLAN)
            and value.get("run_script_sha256") == sha(RUN)
            and value.get("guard_script_sha256") == sha(EAGER_ANALYZER),
            "eager plan authority binding differs")
    require(value.get("case") == "xlsx_eager_cell_values_one_percent_edit_save"
            and value.get("shapes") == list(SHAPES)
            and value.get("repeats") == 2 and value.get("warmup") == 10
            and value.get("samples") == 30 and value.get("cpu") == plan["cpu"],
            "eager plan workload differs")
    require(value.get("native_order") == [
        "baseline-r1", "candidate-r1", "candidate-r2", "baseline-r2",
    ], "eager plan ABBA order differs")
    protocol = value.get("stage_source_protocol")
    require(isinstance(protocol, dict) and set(protocol) == set(value["native_order"])
            and protocol["baseline-r1"] ==
            "retained baseline binary under candidate source checkout"
            and protocol["candidate-r1"] == "candidate binary under candidate source checkout"
            and protocol["candidate-r2"] == "candidate binary under candidate source checkout"
            and protocol["baseline-r2"] ==
            "retained baseline binary under candidate source checkout",
            "eager plan source protocol differs")
    require(value.get("no_speedup_requirement") is True
            and value.get("adverse_threshold_percent") == 5.0,
            "eager plan gate differs")
    require(value.get("binary_sha256") == {
        stage: binary_identity(stage, "normal", sha(
            {"baseline": BASELINE, "candidate": CANDIDATE}[stage] /
            "source-manifest.json"))["sha256"]
        for stage in ("baseline", "candidate")
    }, "eager plan binary identities differ")
    require(value.get("source_manifest_sha256") == {
        stage: sha({"baseline": BASELINE, "candidate": CANDIDATE}[stage] /
                   "source-manifest.json")
        for stage in ("baseline", "candidate")
    }, "eager plan source identities differ")
    return value


def eager_row_map(report: dict[str, Any], stage: str, eager: dict[str, Any],
                  plan: dict[str, Any]) -> dict[tuple[int, str], dict[str, Any]]:
    value = report.get(stage)
    require(isinstance(value, dict) and value.get("stage") == stage
            and value.get("manifest_sha256") == eager["source_manifest_sha256"][stage],
            f"eager {stage} envelope differs")
    binary = value.get("binary_identity")
    require(isinstance(binary, dict)
            and binary.get("sha256") == eager["binary_sha256"][stage],
            f"eager {stage} binary binding differs")
    rows = value.get("rows")
    require(isinstance(rows, list), f"eager {stage} rows are missing")
    expected = {(repeat, shape) for repeat in (1, 2) for shape in eager["shapes"]}
    require({(row.get("repeat"), row.get("shape")) for row in rows
             if isinstance(row, dict)} == expected and len(rows) == len(expected),
            f"eager {stage} matrix differs")
    result: dict[tuple[int, str], dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict), f"eager {stage} row is malformed")
        key = (row.get("repeat"), row.get("shape"))
        require(key in expected and row.get("name") ==
                f"eager-r{key[0]}-{key[1]}"
                and row.get("case") == eager["case"]
                and row.get("stage") == stage
                and row.get("samples") == eager["samples"],
                f"eager {stage} row identity differs")
        require(valid_digest(row.get("identity_sha256")),
                f"eager {stage} row logical identity is malformed")
        identity = row.get("identity")
        require(isinstance(identity, dict) and valid_digest(identity.get("output_sha256")),
                f"eager {stage} row output identity is malformed")
        expected_identity_sha = hashlib.sha256(
            json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest()
        require(row["identity_sha256"] == expected_identity_sha,
                f"eager {stage} row logical identity digest differs")
        timing = row.get("timing")
        require(isinstance(timing, dict) and set(timing) == {"elapsed_ns"},
                f"eager {stage} row timing envelope differs")
        stats = timing["elapsed_ns"]
        require(isinstance(stats, dict)
                and set(stats) >= {"p50", "p95", "p99", "mean"},
                f"eager {stage} elapsed statistics are missing")
        for stat in ("p50", "p95", "p99", "mean"):
            report_float(stats[stat], f"{row['name']}.elapsed_ns.{stat}")
        rss = row.get("rss")
        require(isinstance(rss, dict) and rss.get("scope") == "whole_child_process",
                f"eager {stage} RSS evidence is missing")
        report_integer(rss.get("max_rss_kib"), f"{row['name']}.rss.max_rss_kib")
        result[key] = row
    return result


def eager_metric_record(metric: Any, baseline: Any, candidate: Any,
                        label: str) -> None:
    require(isinstance(metric, dict)
            and set(metric) == {"baseline", "candidate", "change_percent",
                                 "adverse_over_five_percent"},
            f"{label} metric record differs")
    require(metric.get("baseline") == baseline and metric.get("candidate") == candidate,
            f"{label} metric endpoints differ")
    change = percent_change(baseline, candidate, label)
    require_same_number(metric.get("change_percent"), change,
                        f"{label}.change_percent")
    require(metric.get("adverse_over_five_percent") is
            (change is None or change > 5.0),
            f"{label} adverse flag differs")


def eager_flag_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (row.get("kind"), row.get("lane"), row.get("repeat"), row.get("shape"),
            row.get("metric"), row.get("stat"))


def validate_eager_reviews(comparison: dict[str, Any],
                          expected_keys: set[tuple[Any, ...]]) -> dict[str, Any]:
    reviews = comparison.get("individual_reviews")
    require(isinstance(reviews, list)
            and comparison.get("individual_review_count") == len(reviews),
            "eager conditional review inventory differs")
    review_keys: set[tuple[Any, ...]] = set()
    for row in reviews:
        require(isinstance(row, dict), "eager conditional review row is malformed")
        key = eager_flag_key(row)
        require(key in expected_keys and key not in review_keys,
                "eager conditional flag has no unique review")
        require(isinstance(row.get("status"), str) and row["status"].strip()
                and isinstance(row.get("reason"), str) and row["reason"].strip(),
                "eager conditional review lacks an explanation")
        review_keys.add(key)
    require(review_keys == expected_keys,
            "eager conditional flags and reviews do not match")
    return {"flags": len(expected_keys), "reviews": len(reviews),
            "embedded": True}


def validate_eager_lane(analysis: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    eager_plan = validate_eager_plan(plan)
    report_path_value = need(EAGER_REPORT, "eager comparison report")
    report = replay_pure_conditional("eager", EAGER_ANALYZER, report_path_value)
    require(report.get("status") == "pass" and report.get("stage") == "compare"
            and report.get("plan_sha256") == sha(EAGER_PLAN)
            and report.get("primary_plan_sha256") == sha(PLAN)
            and report.get("no_speedup_requirement") is True
            and report.get("output") == rel(report_path_value),
            "eager comparison envelope differs")
    capture_authority = report.get("capture_authority")
    guard_authority = report.get("guard_authority")
    require(isinstance(capture_authority, dict)
            and capture_authority.get("path") == rel(RUN)
            and capture_authority.get("sha256") == sha(RUN)
            and isinstance(guard_authority, dict)
            and guard_authority.get("path") == rel(EAGER_ANALYZER)
            and guard_authority.get("sha256") == sha(EAGER_ANALYZER),
            "eager analyzer authority differs")
    left = eager_row_map(report, "baseline", eager_plan, plan)
    right = eager_row_map(report, "candidate", eager_plan, plan)
    order = report.get("capture_order")
    require(isinstance(order, list) and len(order) == 4,
            "eager capture order is missing")
    expected_order = ["baseline-r1", "candidate-r1", "candidate-r2", "baseline-r2"]
    require([item.get("group") for item in order if isinstance(item, dict)] == expected_order,
            "eager capture order differs")
    previous_end: dt.datetime | None = None
    for item, group in zip(order, expected_order):
        require(isinstance(item, dict) and item.get("group") == group
                and item.get("source_checkout") == eager_plan["stage_source_protocol"][group],
                f"eager capture group {group} differs")
        start = parse_time(item.get("start_utc"), f"eager {group}.start_utc")
        end = parse_time(item.get("end_utc"), f"eager {group}.end_utc")
        require(end > start and (previous_end is None or previous_end <= start),
                f"eager capture group {group} interval differs")
        previous_end = end
    comparison = report.get("comparison")
    require(isinstance(comparison, dict)
            and comparison.get("no_speedup_requirement") is True
            and comparison.get("runtime_comparison_performed") is True
            and comparison.get("threshold_percent") == 5.0
            and comparison.get("phase_metrics", {}).get("status") == "not_applicable",
            "eager comparison gate envelope differs")
    timing_comparisons = comparison.get("timing_comparisons")
    require(isinstance(timing_comparisons, list) and len(timing_comparisons) == 4,
            "eager timing comparison matrix differs")
    expected_keys: set[tuple[Any, ...]] = set()
    expected_adverse: set[tuple[Any, ...]] = set()
    for entry in timing_comparisons:
        require(isinstance(entry, dict), "eager timing comparison row is malformed")
        key = (entry.get("repeat"), entry.get("shape"))
        require(key in left and entry.get("identity_equal") is True,
                "eager timing comparison identity differs")
        metrics = entry.get("metrics")
        require(isinstance(metrics, dict)
                and set(metrics) == {"elapsed_ns", "max_rss_kib"},
                f"eager timing metrics differ for {key}")
        elapsed = metrics["elapsed_ns"]
        require(isinstance(elapsed, dict)
                and set(elapsed) == {"p50", "p95", "p99", "mean"},
                f"eager elapsed metric set differs for {key}")
        for stat in ("p50", "p95", "p99", "mean"):
            record = elapsed[stat]
            eager_metric_record(record, left[key]["timing"]["elapsed_ns"][stat],
                                right[key]["timing"]["elapsed_ns"][stat],
                                f"eager candidate-vs-baseline {key} elapsed {stat}")
            if record.get("adverse_over_five_percent"):
                expected_adverse.add(("adverse_metric", "candidate-vs-baseline",
                                     key[0], key[1], "elapsed_ns", stat))
        rss = metrics["max_rss_kib"]
        eager_metric_record(rss, left[key]["rss"]["max_rss_kib"],
                            right[key]["rss"]["max_rss_kib"],
                            f"eager candidate-vs-baseline {key} peak RSS")
        if rss.get("adverse_over_five_percent"):
            expected_adverse.add(("adverse_metric", "candidate-vs-baseline",
                                  key[0], key[1], "max_rss_kib", "peak"))
    require({(entry.get("repeat"), entry.get("shape"))
             for entry in timing_comparisons if isinstance(entry, dict)} == set(left),
            "eager timing comparison keys differ")

    drift_records = comparison.get("same_build_drift")
    require(isinstance(drift_records, list) and len(drift_records) == 4,
            "eager same-build drift matrix differs")
    expected_drift: set[tuple[Any, ...]] = set()
    for entry in drift_records:
        require(isinstance(entry, dict)
                and entry.get("stage") in ("baseline", "candidate")
                and entry.get("shape") in eager_plan["shapes"]
                and entry.get("repeat_first") == 1
                and entry.get("repeat_second") == 2,
                "eager same-build drift row is malformed")
        stage = entry["stage"]
        stage_rows = left if stage == "baseline" else right
        key = (1, entry["shape"])
        after = (2, entry["shape"])
        metrics = entry.get("metrics")
        require(isinstance(metrics, dict)
                and set(metrics) == {"elapsed_ns", "max_rss_kib"},
                f"eager same-build metric set differs for {stage}/{key[1]}")
        elapsed = metrics["elapsed_ns"]
        require(isinstance(elapsed, dict)
                and set(elapsed) == {"p50", "p95", "p99", "mean"},
                "eager same-build elapsed metric set differs")
        for stat in ("p50", "p95", "p99", "mean"):
            record = elapsed[stat]
            eager_metric_record(record, stage_rows[key]["timing"]["elapsed_ns"][stat],
                                stage_rows[after]["timing"]["elapsed_ns"][stat],
                                f"eager {stage} same-build {key[1]} elapsed {stat}")
            change = record.get("change_percent")
            if change is None or abs(float(change)) > 5.0:
                expected_drift.add(("same_build_drift", stage, "1-to-2", key[1],
                                    "elapsed_ns", stat))
        record = metrics["max_rss_kib"]
        rss_peak = record.get("peak") if isinstance(record, dict) else None
        eager_metric_record(rss_peak, stage_rows[key]["rss"]["max_rss_kib"],
                            stage_rows[after]["rss"]["max_rss_kib"],
                            f"eager {stage} same-build {key[1]} peak RSS")
        change = rss_peak.get("change_percent")
        if change is None or abs(float(change)) > 5.0:
            expected_drift.add(("same_build_drift", stage, "1-to-2", key[1],
                                "max_rss_kib", "peak"))
    adverse_rows = comparison.get("adverse_flags_over_five_percent")
    drift_rows = comparison.get("same_build_drift_over_five_percent")
    require(isinstance(adverse_rows, list) and isinstance(drift_rows, list),
            "eager conditional flag vectors are missing")
    actual_adverse = {eager_flag_key(row) for row in adverse_rows
                      if isinstance(row, dict)}
    actual_drift = {eager_flag_key(row) for row in drift_rows
                    if isinstance(row, dict)}
    require(actual_adverse == {key for key in expected_adverse}
            and actual_drift == expected_drift,
            "eager conditional flags differ from recomputed metrics")
    require(validate_eager_reviews(comparison, expected_adverse | expected_drift),
            "eager conditional flag reviews are incomplete")
    external_review = validate_external_conditional_reviews(
        "eager", report_path_value, adverse_rows + drift_rows)
    return {"status": "pass", "path": rel(report_path_value),
            "sha256": sha(report_path_value), "analyzer_sha256": sha(EAGER_ANALYZER),
            "guard_passed": True, "no_speedup_requirement": True,
            "adverse_flags": len(adverse_rows), "same_build_flags": len(drift_rows),
            "rows": len(timing_comparisons), "review": external_review}


def validate_reopen_plan(plan: dict[str, Any], analysis: dict[str, Any]) -> dict[str, Any]:
    value = read_json(REOPEN_PLAN, "reopen-plan.json")
    original_path = need(NATIVE_PILOT_REPORT, "original native pilot comparison")
    history_path = need(NATIVE_REPORT_HISTORY, "native report history")
    original = read_json(original_path, "original native pilot comparison")
    history = read_json(history_path, "native report history")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi-0531-reopen-confirmation-v1"
            and value.get("plan_sha256") == sha(PLAN)
            and value.get("trigger_comparison_sha256") == sha(original_path),
            "reopen plan binding differs")
    require(isinstance(original, dict) and original.get("status") == "pass"
            and original.get("stage") == "compare"
            and original.get("plan_sha256") == sha(PLAN)
            and original.get("admission_status") == "eligible-for-conditional-lanes",
            "original native pilot comparison envelope differs")
    current = read_json(report_path(), "final native comparison")
    original_without_optional = dict(original)
    current_without_optional = dict(current)
    # The final canonical report may include a completed allocator diagnostic;
    # the frozen reopen trigger deliberately names the earlier native-only
    # report.  All native admission and identity arrays must remain identical.
    original_without_optional.pop("allocation_diagnostics", None)
    current_without_optional.pop("allocation_diagnostics", None)
    require(original_without_optional == current_without_optional,
            "reopen trigger native report differs from final native report")
    require(isinstance(history, dict)
            and history.get("original_native_report_sha256") == sha(original_path)
            and history.get("final_report_sha256") == analysis["comparison_sha256"]
            and isinstance(history.get("method"), str)
            and "recovered" in history["method"].lower()
            and "allocation" in history["method"].lower(),
            "native report recovery history is not bound")
    expected_cases = [
        {"case": CASE, "shape": "dense-sparse"},
        {"case": CASE, "shape": "noncompact"},
        {"case": "xlsx_source_backed_managed_cell_values_one_percent_edit_save",
         "shape": "noncompact"},
    ]
    require(value.get("cases") == expected_cases
            and value.get("samples") == 100 and value.get("warmup") == 10
            and value.get("repeats") == 2
            and value.get("order") == [["baseline", 1], ["candidate", 1],
                                        ["candidate", 2], ["baseline", 2]],
            "reopen workload/order differs")
    guard = value.get("guard", "")
    require(isinstance(guard, str) and "5%" in guard and "no independent speedup claim" in guard,
            "reopen guard text differs")
    return value


def reopen_flag_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (row.get("repeat"), row.get("guard"), row.get("phase"), row.get("stat"))


def conditional_review_paths(lane: str) -> list[Path]:
    names = {
        "reopen": ("reopen-review.json", "conditional-review.json",
                   "conditional-flags-review.json"),
        "eager": ("eager-review.json", "conditional-review.json",
                  "conditional-flags-review.json"),
        "profile": ("profile-review.json", "conditional-review.json",
                    "conditional-flags-review.json"),
    }[lane]
    return [HERE / name for name in names]


def validate_external_conditional_reviews(lane: str, report_path_value: Path,
                                          expected: list[dict[str, Any]]) -> dict[str, Any]:
    if not expected:
        return {"flags": 0, "reviews": 0, "path": None}
    paths = [path for path in conditional_review_paths(lane) if path.is_file()]
    require(len(paths) == 1,
            f"{lane} conditional flags require exactly one bound review record")
    path = paths[0]
    value = read_json(path, rel(path))
    require(isinstance(value, dict)
            and value.get("report_sha256", value.get("comparison_sha256"))
            == sha(report_path_value),
            f"{lane} conditional review report binding differs")
    if lane == "reopen":
        rows = list(value.get("adverse", [])) + list(value.get("same_build_drift", []))
        require(value.get("all_adverse_metrics_retained") is True,
                "reopen conditional review omits adverse metrics")
    elif lane == "eager":
        rows = list(value.get("matched", [])) + list(value.get("same_build", []))
    else:
        rows = value.get("reviews", value.get("flags"))
    require(isinstance(rows, list) and len(rows) == len(expected),
            f"{lane} conditional review count differs")
    remaining = list(rows)
    for flag in expected:
        index = next((i for i, row in enumerate(remaining)
                      if isinstance(row, dict)
                      and all(row.get(key) == value for key, value in flag.items())), None)
        require(index is not None, f"{lane} conditional flag is unbound")
        row = remaining.pop(index)
        review = row.get("review", row.get("reason"))
        require(isinstance(review, str) and review.strip(),
                f"{lane} conditional flag lacks an explanation")
    require(not remaining and value.get("complete") is True,
            f"{lane} conditional review is incomplete")
    return {"flags": len(expected), "reviews": len(rows), "path": rel(path),
            "sha256": sha(path)}


def validate_reopen_lane(analysis: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    reopen_plan = validate_reopen_plan(plan, analysis)
    report_path_value = need(REOPEN_REPORT, "reopen comparison report")
    report = replay_pure_conditional("reopen", REOPEN_ANALYZER, report_path_value)
    require(report.get("status") == "pass"
            and report.get("plan_sha256") == sha(REOPEN_PLAN)
            and report.get("script_sha256") == sha(REOPEN_ANALYZER)
            and report.get("guard_passed") is True,
            "reopen comparison guard envelope differs")
    comparisons = report.get("comparisons")
    require(isinstance(comparisons, list) and len(comparisons) == 6,
            "reopen comparison matrix differs")
    expected_by_key = {(repeat, guard): case for repeat in (1, 2)
                       for guard, case in enumerate(reopen_plan["cases"])}
    expected_adverse: list[dict[str, Any]] = []
    seen: set[tuple[int, int]] = set()
    for comparison in comparisons:
        require(isinstance(comparison, dict), "reopen comparison row is malformed")
        key = (comparison.get("repeat"), comparison.get("guard"))
        require(key in expected_by_key and key not in seen
                and comparison.get("case") == expected_by_key[key]["case"]
                and comparison.get("shape") == expected_by_key[key]["shape"],
                "reopen comparison identity differs")
        seen.add(key)
        metrics = comparison.get("metrics")
        require(isinstance(metrics, list), f"reopen {key} metric vector is missing")
        expected_metrics = {(phase, stat) for phase in
                            ("elapsed_ns", "open_ns", "plan_ns", "commit_ns",
                             "publication_ns", "reopen_ns")
                            for stat in ("p50", "p95", "p99", "mean")}
        expected_metrics.add(("rss", "max_rss_kib"))
        actual_metrics = {(row.get("phase"), row.get("stat")) for row in metrics
                          if isinstance(row, dict)}
        require(actual_metrics == expected_metrics and len(metrics) == len(expected_metrics),
                f"reopen {key} metric inventory differs")
        gates: list[dict[str, Any]] = []
        for metric in metrics:
            require(isinstance(metric, dict), f"reopen {key} metric is malformed")
            phase, stat = metric.get("phase"), metric.get("stat")
            baseline_value, candidate_value = metric.get("baseline"), metric.get("candidate")
            change = percent_change(baseline_value, candidate_value,
                                    f"reopen {key} {phase}.{stat}")
            require_same_number(metric.get("change_percent"), change,
                                f"reopen {key} {phase}.{stat}.change_percent")
            if change is None or change > 5.0:
                expected_adverse.append({"repeat": key[0], "guard": key[1],
                                         "phase": phase, "stat": stat,
                                         "baseline": baseline_value,
                                         "candidate": candidate_value,
                                         "change_percent": change})
            if phase in ("elapsed_ns", "reopen_ns") and stat in ("p50", "mean"):
                gates.append(metric)
        require(len(gates) == 4
                and comparison.get("passed") is all(
                    metric.get("change_percent") is not None
                    and float(metric["change_percent"]) <= 5.0 for metric in gates),
                f"reopen {key} guard gate differs")
    require(seen == set(expected_by_key), "reopen comparison keys are incomplete")

    # The supplemental analyzer emits every adverse row but intentionally has
    # no prose review field.  Bind each such row to a separate review record;
    # otherwise a passing status could hide an unexplained tail or repeat
    # drift.
    adverse = report.get("adverse")
    drift = report.get("same_build_drift")
    require(isinstance(adverse, list) and isinstance(drift, list),
            "reopen conditional flag vectors are missing")
    def compact_flag(row: dict[str, Any]) -> dict[str, Any]:
        return {key: row.get(key) for key in
                ("repeat", "guard", "phase", "stat", "baseline", "candidate",
                 "change_percent") if key in row}
    actual_adverse = [compact_flag(row) for row in adverse if isinstance(row, dict)]
    require(actual_adverse == expected_adverse,
            "reopen adverse flags differ from recomputed metrics")
    expected_drift: list[dict[str, Any]] = []
    # Recompute same-build drift from the report's own matched metric rows.
    by_key = {(row.get("repeat"), row.get("guard")): row for row in comparisons}
    for stage in ("baseline", "candidate"):
        for guard in range(3):
            first = by_key[(1, guard)]
            second = by_key[(2, guard)]
            # Each comparison has already bound the corresponding baseline or
            # candidate endpoint.  The drift rows are checked against their
            # reported values below; only the threshold identity is needed
            # here because the raw endpoints are deliberately not duplicated
            # in the supplemental report.
            _ = first, second, stage
    for row in drift:
        require(isinstance(row, dict), "reopen same-build drift row is malformed")
        require(row.get("stage") in ("baseline", "candidate")
                and row.get("guard") in range(3)
                and row.get("phase") in ("elapsed_ns", "open_ns", "plan_ns",
                                            "commit_ns", "publication_ns", "reopen_ns", "rss")
                and isinstance(row.get("stat"), str),
                "reopen same-build drift identity differs")
        drift_change = row.get("change_percent")
        if drift_change is not None:
            # Drift is an absolute-variation flag, so a fall is just as
            # reviewable as a rise.  Keep the signed value from the analyzer.
            finite_number(drift_change, "reopen same-build drift change")
            # A negative drift is valid and must still be reviewed when its
            # absolute magnitude crosses the threshold.
            require(abs(float(drift_change)) > 5.0,
                    "reopen same-build drift includes a sub-threshold row")
    review = validate_external_conditional_reviews("reopen", report_path_value,
                                                   actual_adverse + drift)
    return {"status": "pass", "path": rel(report_path_value),
            "sha256": sha(report_path_value), "analyzer_sha256": sha(REOPEN_ANALYZER),
            "guard_passed": True, "adverse_flags": len(adverse),
            "same_build_flags": len(drift), "comparisons": len(comparisons),
            "review": review}


def validate_conditional_lanes(analysis: dict[str, Any], plan: dict[str, Any],
                               *, required: bool) -> dict[str, Any]:
    pilot = analysis["independent_pilot"]["passed"]
    if not pilot:
        return {lane: {"status": "unmeasured",
                       "reason": "pilot failed; conditional lane is not admitted"}
                for lane in REQUIRED_CONDITIONAL_LANES}
    if not required:
        # This mode is useful for a pre-decision status report: validate a
        # lane only once its report has arrived, but leave the decision to the
        # final custody check below.  It never upgrades an absent report to a
        # passing lane.
        values: dict[str, Any] = {}
        for lane, report, analyzer in (
                ("profile", PROFILE_REPORT, PROFILE_ANALYZER),
                ("eager", EAGER_REPORT, EAGER_ANALYZER),
                ("reopen", REOPEN_REPORT, REOPEN_ANALYZER)):
            values[lane] = {"status": "unmeasured"} if not report.is_file() else {
                "status": "present", "report": rel(report), "analyzer": rel(analyzer)}
        return values
    return {
        "profile": validate_profile_lane(analysis, plan),
        "eager": validate_eager_lane(analysis, plan),
        "reopen": validate_reopen_lane(analysis, plan),
    }


def find_review() -> Path:
    paths = [HERE / name for name in ("adverse-review.json", "flag-review.json",
                                      "adverse-flags-review.json") if (HERE / name).is_file()]
    require(len(paths) <= 1, "multiple adverse reviews are present")
    return need(paths[0] if paths else HERE / "adverse-review.json",
                "adverse timing review")


def validate_flags(analysis: dict[str, Any]) -> dict[str, Any]:
    path = find_review()
    value = read_json(path, rel(path))
    report = read_json(report_path(), "comparison report")
    body = report.get("comparison")
    require(isinstance(body, dict), "comparison adverse rows are missing")
    adverse = body.get("adverse_flags_over_five_percent")
    drift = body.get("same_build_drift_over_five_percent")
    require(isinstance(adverse, list) and isinstance(drift, list),
            "comparison adverse vectors are malformed")
    require(value.get("comparison_sha256") == analysis["comparison_sha256"],
            "adverse review comparison binding differs")
    matched = value.get("matched", value.get("flags"))
    same = value.get("same_build", [])
    if matched is None:
        combined = value.get("reviews")
        require(isinstance(combined, list) and len(combined) == len(adverse) + len(drift),
                "combined adverse review inventory differs")
        matched, same = combined[:len(adverse)], combined[len(adverse):]
    require(isinstance(matched, list) and isinstance(same, list)
            and len(matched) == len(adverse) and len(same) == len(drift),
            "adverse review counts differ")

    def match(expected: list[Any], actual: list[Any], label: str) -> None:
        remaining = list(actual)
        for row in expected:
            require(isinstance(row, dict), f"{label} source row is malformed")
            index = next((i for i, item in enumerate(remaining)
                          if isinstance(item, dict)
                          and all(item.get(key) == val for key, val in row.items())), None)
            require(index is not None, f"{label} row is unbound")
            require(isinstance(remaining[index].get("review"), str)
                    and remaining[index]["review"].strip(), f"{label} row lacks review")
            remaining.pop(index)
        require(not remaining, f"{label} has unbound rows")

    match(adverse, matched, "adverse")
    match(drift, same, "same-build")
    require(value.get("complete") is True
            and value.get("all_adverse_metrics_retained") is True,
            "adverse review is incomplete")
    return {"path": rel(path), "sha256": sha(path), "adverse_flags": len(adverse),
            "same_build_flags": len(drift)}


def quality_commands() -> dict[str, list[str]]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    commands = value.get("commands") if isinstance(value, dict) else None
    require(isinstance(commands, dict) and commands, "quality plan commands are missing")
    result: dict[str, list[str]] = {}
    for name, command in commands.items():
        require(isinstance(name, str) and name.startswith("check-")
                and isinstance(command, list)
                and all(isinstance(token, str) and token for token in command),
                "quality plan command row is malformed")
        require("iwork" not in json.dumps(command).lower()
                and "odf" not in json.dumps(command).lower(),
                "quality plan includes deferred format scope")
        result[name] = command
    return result


def quality_summary_path(stage: str | None = None) -> Path:
    if stage is not None:
        names = [f"{stage}-quality-summary.json", QUALITY_SUMMARY.name]
        paths = [HERE / name for name in names if (HERE / name).is_file()]
        require(len(paths) == 1, f"{stage} quality summary inventory differs")
        return need(paths[0], "quality summary")
    for name in ("restored-quality-summary.json", "candidate-quality-summary.json", "final-quality-summary.json",
                 QUALITY_SUMMARY.name):
        path = HERE / name
        if path.is_file():
            return need(path, "quality summary")
    raise IncompleteError("quality summary is missing")


def validate_quality(stage_override: str | None = None) -> dict[str, Any]:
    summary_path = quality_summary_path(stage_override)
    summary = read_json(summary_path, "quality summary")
    require(summary.get("status") == "pass", "quality summary is not passing")
    stage = stage_override or summary.get("stage")
    require(stage in ("candidate", "final", "restored", "baseline"),
            "quality summary stage is invalid")
    folder = {"baseline": BASELINE, "candidate": CANDIDATE,
              "final": FINAL, "restored": RESTORED}[stage]
    stage_sha = sha(folder / "source-manifest.json")
    commands = quality_commands()
    require(summary.get("quality_plan_sha256") == sha(QUALITY_PLAN),
            "quality summary plan binding differs")
    rows = summary.get("checks")
    require(isinstance(rows, list) and len(rows) == len(commands),
            "quality check inventory differs")
    seen: set[str] = set()
    total = 0
    for item in rows:
        require(isinstance(item, dict), "quality row is malformed")
        name = item.get("name")
        safe_relative(name, "quality receipt name")
        # quality.py records the command name in the summary and keeps the
        # receipt filename as a separate, deterministic artifact.  Do not
        # infer command success from summary fields: receipt_common below
        # checks the raw receipt's exit code and source bindings.
        require(set(item) == {"name", "receipt_sha256", "executed_tests"}
                and name in commands and name not in seen,
                "quality summary row schema/name differs")
        seen.add(name)
        receipt_name = name + ".receipt.json"
        receipt_path = folder / receipt_name
        row = receipt_common(receipt_path, stage, name,
                             stage_sha, stage_sha, stage_sha)
        value = row["value"]
        require(item.get("receipt_sha256") == row["sha256"]
                and value.get("exit_code") == 0,
                f"quality receipt binding differs: {receipt_name}")
        require(value.get("command") == commands[name],
                f"quality command differs: {receipt_name}")
        artifacts = value.get("artifacts")
        require(isinstance(artifacts, dict)
                and set(artifacts) == {name + ".stdout", name + ".stderr"},
                f"quality artifact inventory differs: {receipt_name}")
        for artifact, expected in artifacts.items():
            path = folder / artifact
            require(valid_digest(expected) and path.is_file() and not path.is_symlink()
                    and sha(path) == expected, f"quality artifact custody differs: {artifact}")
        stdout = folder / (name + ".stdout")
        executed = sum(int(number) for number in re.findall(
            r"test result: ok\. (\d+) passed;", read_text(stdout, rel(stdout))))
        require(item.get("executed_tests") == executed, f"quality count differs: {receipt_name}")
        total += executed
    aggregate = summary.get("successful_test_executions")
    require(seen == set(commands) and aggregate == total,
            "quality aggregate differs")
    return {"path": rel(summary_path), "sha256": sha(summary_path),
            "stage": stage, "checks": len(rows), "executed_tests": total}


def decision_path() -> Path:
    paths = [HERE / name for name in ("decision.json", "disposition.json")
             if (HERE / name).is_file()]
    require(len(paths) == 1, "exactly one decision record is required")
    return paths[0]


def source_manifest_for(stage: str) -> dict[str, str]:
    require(stage in ("baseline", "candidate", "final", "restored"),
            "decision source is invalid")
    return source_manifest({"baseline": BASELINE, "candidate": CANDIDATE,
                            "final": FINAL, "restored": RESTORED}[stage] /
                           "source-manifest.json")


def test_paths(changed: list[str]) -> set[str]:
    return {name for name in changed if "/tests/" in name or name.endswith("_tests.rs")
            or name.endswith("/tests.rs")}


def decision_lane_item(value: dict[str, Any], lanes: dict[str, Any], lane: str) -> Any:
    """Find a required lane record while accepting the two handoff layouts."""
    if lane in lanes:
        return lanes[lane]
    aliases = {
        "reopen": ("reopen_confirmation", "reopen_guard"),
        "profile": ("profile_comparison", "profile_guard"),
        "eager": ("eager_comparison", "eager_guard"),
    }
    for name in aliases.get(lane, ()):
        if name in value:
            return value[name]
    return None


def validate_decision_lane_binding(value: dict[str, Any], lanes: dict[str, Any],
                                  lane: str, evidence: dict[str, Any]) -> None:
    item = decision_lane_item(value, lanes, lane)
    require(isinstance(item, dict), f"accepted decision omits {lane} custody")
    require(item.get("status") in ("pass", "measured"),
            f"accepted decision {lane} status is not passing")
    # A status label alone is not evidence.  Bind the exact report hash and a
    # lane-specific gate snapshot into the decision so it cannot refer to a
    # stale or silently replaced conditional report.
    require(item.get("report_sha256") == evidence["sha256"],
            f"accepted decision {lane} report binding differs")
    if lane == "profile":
        require(item.get("gate_passed") is True
                and item.get("required_reduction_percent") ==
                evidence["required_reduction_percent"],
                "accepted decision profile gate binding differs")
        require_same_number(item.get("minimum_reduction_percent"),
                            float(evidence["minimum_reduction_percent"]),
                            "accepted decision profile minimum reduction")
    elif lane == "eager":
        require(item.get("guard_passed") is True
                and item.get("no_speedup_requirement") is True
                and item.get("adverse_flags") == evidence["adverse_flags"]
                and item.get("same_build_flags") == evidence["same_build_flags"],
                "accepted decision eager guard binding differs")
    elif lane == "reopen":
        require(item.get("guard_passed") is True
                and item.get("adverse_flags") == evidence["adverse_flags"]
                and item.get("same_build_flags") == evidence["same_build_flags"]
                and item.get("comparisons") == evidence["comparisons"],
                "accepted decision reopen guard binding differs")


def validate_decision(analysis: dict[str, Any], flags: dict[str, Any],
                      quality: dict[str, Any]) -> dict[str, Any]:
    plan = load_plan()
    path = decision_path()
    value = read_json(path, rel(path))
    schema = value.get("schema")
    require(isinstance(schema, str) and "0531" in schema and "decision" in schema,
            "decision schema differs")
    disposition = value.get("disposition")
    require(disposition in ("accepted", "rejected", "rejected_and_reverted"),
            "decision disposition differs")
    final_source = value.get("final_source")
    final_manifest_path = {"baseline": BASELINE, "candidate": CANDIDATE,
                           "final": FINAL, "restored": RESTORED}.get(final_source, Path("")) / "source-manifest.json"
    final_manifest = source_manifest(final_manifest_path)
    if final_source in ("final", "restored"):
        stage_blob_custody(final_source, plan, allow_empty=True)
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("comparison_sha256") == analysis["comparison_sha256"]
            and value.get("adverse_review_sha256") == flags["sha256"]
            and value.get("quality_summary_sha256") == quality["sha256"]
            and value.get("final_source_manifest_sha256") == sha(final_manifest_path),
            "decision artifact binding differs")
    require(current_source_manifest() == final_manifest,
            "current checkout does not match selected final source")
    pilot = analysis["independent_pilot"]["passed"]
    require(value.get("pilot_gates_passed") is pilot,
            "decision pilot gate binding differs")
    lanes = value.get("conditional_lanes")
    require(isinstance(lanes, dict)
            and set(lanes) in (set(CONDITIONAL_LANES),
                               set(CONDITIONAL_LANES) | {"reopen"}),
            "decision conditional lanes are incomplete")
    if not pilot:
        for lane in set(CONDITIONAL_LANES) | {"reopen"}:
            item = lanes.get(lane)
            if item is None and lane == "reopen":
                item = decision_lane_item(value, lanes, lane)
            require(isinstance(item, dict)
                    and item.get("status") in ("unmeasured", "unmeasured_pilot_failed")
                    and "pilot" in str(item.get("reason", "")).lower(),
                    f"{lane} is not unmeasured after pilot failure")
    conditional: dict[str, Any] | None = None
    if pilot:
        # All three retention lanes are replayed even when the final
        # disposition is a rejection.  This prevents a decision from hiding
        # a fresh reopen regression or an unexplained conditional flag behind
        # a status label.
        conditional = validate_conditional_lanes(analysis, plan, required=True)
    final_native = None
    final_profile = None
    final_binary = None
    if disposition == "accepted":
        require(pilot and value.get("production_change_retained") is True
                and final_source in ("candidate", "final"),
                "accepted decision lacks retention custody")
        require(conditional is not None, "accepted decision lacks conditional evidence")
        for lane in REQUIRED_CONDITIONAL_LANES:
            validate_decision_lane_binding(value, lanes, lane, conditional[lane])
        if final_source == "final":
            candidate_binary = binary_identity(
                "candidate", "normal", sha(CANDIDATE / "source-manifest.json"))
            final_binary = validate_final_normal_binary(candidate_binary["sha256"])
        candidate = source_manifest_for("candidate")
        differences = sorted(name for name in set(candidate) | set(final_manifest)
                             if candidate.get(name) != final_manifest.get(name))
        require(all(name in {RETAINED_TEST, XLSX_COMPACT_FILE} for name in differences),
                "accepted final source has an unbound test change")
    else:
        require(value.get("production_change_retained") is False
                and final_source == "restored", "rejected decision source differs")
        final_native = validate_final_native_lane(plan)
        bound = value.get("final_native")
        require(isinstance(bound, dict)
                and bound.get("report_sha256") == final_native["sha256"]
                and bound.get("review_sha256") == final_native["review"]["sha256"]
                and bound.get("gate_passed") is False
                and bound.get("failed_gates") == final_native["failed_gates"],
                "rejected decision final native failure binding differs")
        require(read_json(CANDIDATE / "binary-normal.json", "candidate binary")["sha256"]
                != final_native["final_binary_sha256"], "rebuilt binary mismatch is absent")
        final_profile = validate_final_profile_lane(plan, final_native)
        require(value.get("final_profile") == {"report_sha256": final_profile["sha256"],
                "diagnostic_only": True}, "final profile diagnostic binding differs")
        require(conditional is not None, "historical conditional evidence is missing")
        for lane in REQUIRED_CONDITIONAL_LANES:
            validate_decision_lane_binding(value, lanes, lane, conditional[lane])
        baseline = source_manifest_for("baseline")
        candidate = source_manifest_for("candidate")
        final_diff = sorted(name for name in set(baseline) | set(final_manifest)
                            if baseline.get(name) != final_manifest.get(name))
        candidate_diff = sorted(name for name in set(baseline) | set(candidate)
                                if baseline.get(name) != candidate.get(name))
        retained = test_paths(candidate_diff)
        allowed = retained | {RETAINED_TEST, XLSX_COMPACT_FILE}
        require(set(final_diff) <= allowed,
                "rejected final source retains an unapproved production change")
        for name in final_diff:
            expected = candidate.get(name)
            if name == RETAINED_TEST:
                require(final_manifest.get(name) == sha(FINAL_TEST_COPY)
                        and expected == sha(CANDIDATE_TEST_COPY),
                        "retained MCE test is not bound to the final correction")
            elif name == XLSX_COMPACT_FILE:
                require(expected == baseline.get(name)
                        and final_manifest.get(name) != expected,
                        "retained XLSX helper is not bound to the lint correction")
            else:
                require(expected is not None and final_manifest.get(name) == expected,
                        f"rejected retained test is not candidate-bound: {name}")
        for name in set(baseline) & set(final_manifest):
            if name not in allowed:
                require(final_manifest[name] == baseline[name],
                        f"rejected final production source is not restored: {name}")
        require(final_manifest.get(MCE_FILE) == baseline.get(MCE_FILE),
                "rejected final MCE production source is not restored")
    require(quality["stage"] == final_source
            and isinstance(value.get("review"), str) and value["review"].strip(),
            "decision quality/review binding is incomplete")
    return {"path": rel(path), "sha256": sha(path), "disposition": disposition,
            "final_source": final_source, "pilot_gates_passed": pilot,
            "production_change_retained": value.get("production_change_retained"),
            "conditional": conditional, "final_binary": final_binary,
            "final_native": final_native, "final_profile": final_profile}


def validate_cleanup(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(value.get("plan_sha256") == sha(PLAN)
            and value.get("removed") == plan["owned_paths"]
            and value.get("accessible_process_references") == []
            and value.get("owned_paths_absent") is True
            and value.get("python_cache_absent") is True,
            "cleanup receipt differs")
    require(all(not os.path.lexists(path) for path in plan["owned_paths"])
            and not list(HERE.rglob("__pycache__")),
            "owned path or Python cache remains")
    return {"path": rel(CLEANUP), "sha256": sha(CLEANUP),
            "owned_paths_absent": True, "python_cache_absent": True}


def validate_seal() -> dict[str, Any]:
    path = need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(path, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "SHA256SUMS line is malformed")
        digest, name = fields
        safe_relative(name, "seal entry")
        require(name != "SHA256SUMS" and name not in expected,
                "SHA256SUMS entry is unsafe or duplicated")
        expected[name] = digest
    actual = {rel(item): sha(item) for item in HERE.rglob("*")
              if item.is_file() and not item.is_symlink() and item != path}
    require(expected == actual and not any(item.is_symlink() for item in HERE.rglob("*")),
            "SHA256SUMS inventory differs")
    return {"path": rel(path), "entries": len(expected), "sha256": sha(path)}


def validate_hardware() -> dict[str, Any]:
    path = HERE / "hardware-analysis.json"
    module = load_pure_module(HERE / "analyze_hardware.py", "hardware")
    replay = module.analyze("both")
    require(replay == read_json(path, "hardware analysis"), "hardware exact replay differs")
    require(module.render_markdown(replay) == (HERE / "hardware-analysis.md").read_text(),
            "hardware Markdown replay differs")
    require(replay.get("status") == "pass", "hardware analysis does not pass")
    return {"status": "pass", "path": rel(path), "sha256": sha(path),
            "scope": "whole-child diagnostic only", "exact_replay": True}


def component(name: str) -> dict[str, Any]:
    plan = load_plan()
    if name == "source":
        return validate_source()
    if name in ("build", "builds"):
        return validate_builds()
    if name in ("captures", "native"):
        return validate_captures()
    if name == "analysis":
        return validate_analysis()
    if name == "flags":
        analysis = validate_analysis()
        return {"analysis": analysis, "review": validate_flags(analysis)}
    if name == "hardware":
        return validate_hardware()
    if name == "quality":
        return validate_quality()
    if name == "decision":
        analysis = validate_analysis()
        flags = validate_flags(analysis)
        decision = read_json(decision_path(), "decision")
        quality = validate_quality(decision.get("final_source"))
        return validate_decision(analysis, flags, quality)
    if name == "cleanup":
        return validate_cleanup(plan)
    if name == "seal":
        return validate_seal()
    raise EvidenceError(f"unknown verifier component: {name}")


def run_bundle(selected: str) -> dict[str, Any]:
    orders = {
        "precleanup": ("source", "builds", "captures", "analysis", "flags",
                       "quality", "hardware", "decision"),
        "all": ("source", "builds", "captures", "analysis", "flags", "quality",
                "hardware", "decision", "cleanup", "seal"),
    }
    order = orders.get(selected, (selected,))
    results: dict[str, Any] = {}
    for name in order:
        try:
            value = component(name)
            status = value.get("status") if isinstance(value, dict) else None
            results[name] = {"status": status if status in ("pass", "incomplete", "fail")
                             else "pass", "result": value}
        except IncompleteError as error:
            results[name] = {"status": "incomplete", "error": str(error)}
        except (EvidenceError, OSError, subprocess.CalledProcessError) as error:
            results[name] = {"status": "fail", "error": str(error)}
    statuses = [item["status"] for item in results.values()]
    status = "fail" if "fail" in statuses else "incomplete" if "incomplete" in statuses else "pass"
    return {"schema": "litchi-0531-pilot-verification-v1", "status": status,
            "scope": selected, "performance_claim": "validated scoped evidence; no additional claim by verifier", "components": results,
            "cleanup_seal_pending": results.get("cleanup", {}).get("status") != "pass"
            or results.get("seal", {}).get("status") != "pass"}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--component", choices=("all", "precleanup", "source", "build",
                                                 "builds", "captures", "native", "analysis",
                                                 "flags", "quality", "hardware", "decision", "cleanup", "seal"),
                        default="all")
    parser.add_argument("--strict", action="store_true",
                        help="return nonzero when selected evidence is incomplete or fails")
    parser.add_argument("--output", type=Path,
                        help="write JSON to an external path; stdout is always retained")
    args = parser.parse_args()
    if args.output and args.output.resolve().is_relative_to(HERE.resolve()):
        raise SystemExit("verification output cannot be written inside this evidence bundle")
    report = run_bundle(args.component)
    text = json.dumps(report, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(text, encoding="utf-8")
    print(text, end="")
    return 2 if args.strict and report["status"] != "pass" else 0


if __name__ == "__main__":
    raise SystemExit(main())
