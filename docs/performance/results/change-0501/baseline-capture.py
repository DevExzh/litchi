#!/usr/bin/env python3
"""Capture and validate two current full default CRUD baseline repeats.

The binary and source hash manifest are prepared before this driver is invoked.
Each lane runs the frozen 37-case/201-row default matrix with one worker, three
warmups, and fifteen retained samples per row.  The report and schema-2 catalog
stay in the lane directory; the catalog sidecar is passed explicitly to the
coverage validator because a current run may have different provenance
metadata from the checked catalog while retaining the same content set.

This driver deliberately does not build, mutate the worktree, or compare the
two repeats.  It records the exact command, binary identity, environment,
resource log, report, catalog, source manifest, and validation result so a
later review can do those jobs from immutable lane artifacts.  The source
manifest's candidate base is the authoritative source identity; the observed
Git HEAD is recorded only as incidental host state.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
CAPTURES = ROOT / "captures"
TEMP_ROOT = Path("/tmp/litchi-goal-0501/default-corpora")
CPU = "2"
WORKERS = "1"
SAMPLES = 15
WARMUPS = 3
EXPECTED_ROWS = 201
EXPECTED_CASES = 37
LANES = ("R1-normal", "R2-normal")
CHECKED_CATALOG = REPO / "docs/performance/results/perf-corpus-manifest-v2.json"
INDEX_VALIDATOR = REPO / "tools/validate_crud_coverage_index.py"
STATIC_TEST_MODULE = "tools.test_crud_coverage_index"


class CaptureError(RuntimeError):
    """A custody, capture, cleanup, or validation failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked JSON: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise CaptureError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"JSON object required: {path}")
    return value


def write_json_exclusive(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"artifact already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2, sort_keys=True)
        stream.write("\n")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(REPO))
    except ValueError:
        return str(path)


def artifact(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def git_revision() -> str:
    result = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=REPO,
        check=True,
        capture_output=True,
        text=True,
    )
    revision = result.stdout.strip()
    require(len(revision) == 40 and all(character in "0123456789abcdef" for character in revision),
            f"unexpected git revision: {revision!r}")
    return revision


def git_status() -> list[str]:
    result = subprocess.run(
        ["git", "status", "--porcelain=v1"],
        cwd=REPO,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.splitlines()


def ensure_source_manifest(value: str) -> Path:
    """Resolve a caller-supplied source manifest without following a file symlink."""

    path = Path(value)
    if not path.is_absolute():
        path = REPO / path
    require(not path.is_symlink(), f"source manifest is symlinked: {path}")
    require(path.is_file(), f"source manifest is missing: {path}")
    return path.resolve()


def load_source_manifest(path: Path) -> dict[str, Any]:
    """Load and immediately verify the final-candidate source hash manifest."""

    raw = read_json(path)
    candidate = next(
        (raw.get(key) for key in ("candidate_base_revision", "candidate_base", "revision")
         if raw.get(key) is not None),
        None,
    )
    require(
        isinstance(candidate, str)
        and len(candidate) == 40
        and candidate == candidate.lower()
        and all(character in "0123456789abcdef" for character in candidate),
        "source manifest candidate base must be a lowercase forty-character Git identity",
    )
    files = raw.get("files")
    require(isinstance(files, dict) and files, "source manifest files map is missing or empty")
    normalized: dict[str, str] = {}
    for name, expected in sorted(files.items()):
        require(isinstance(name, str) and name, "source manifest contains an invalid path")
        relative = Path(name)
        require(not relative.is_absolute() and ".." not in relative.parts,
                f"source manifest path escapes repository: {name!r}")
        require(
            isinstance(expected, str)
            and len(expected) == 64
            and expected == expected.lower()
            and all(character in "0123456789abcdef" for character in expected),
            f"source manifest hash is invalid: {name!r}",
        )
        source = REPO / relative
        require(source.is_file() and not source.is_symlink(),
                f"source manifest file is missing or symlinked: {name}")
        observed = sha(source)
        require(observed == expected, f"source manifest hash differs before capture: {name}")
        normalized[name] = expected
    record = {
        "path": str(path),
        "artifact": artifact(path),
        "candidate_base_revision": candidate,
        "files": normalized,
        "files_count": len(normalized),
    }
    return record


def verify_source_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    """Rehash the manifest and every bound source file, returning a receipt view."""

    manifest_path = Path(str(manifest["path"]))
    expected_artifact = manifest["artifact"]
    current_artifact = artifact(manifest_path)
    require(current_artifact["sha256"] == expected_artifact["sha256"],
            "source manifest changed during capture")
    expected_files = manifest["files"]
    require(isinstance(expected_files, dict) and expected_files,
            "source manifest file map is missing during verification")
    observed_files: dict[str, str] = {}
    for name, expected in sorted(expected_files.items()):
        source = REPO / Path(name)
        require(source.is_file() and not source.is_symlink(),
                f"source manifest file disappeared or became symlinked: {name}")
        observed = sha(source)
        require(observed == expected, f"source changed during capture: {name}")
        observed_files[name] = observed
    return {"manifest": current_artifact, "files": observed_files}


def ensure_binary(value: str) -> Path:
    path = Path(value)
    require(path.is_absolute(), "binary path must be absolute")
    require(not path.is_symlink(), f"binary is symlinked: {path}")
    path = path.resolve(strict=False)
    require(path.is_file() and not path.is_symlink(), f"binary is missing or symlinked: {path}")
    require(os.access(path, os.X_OK), f"binary is not executable: {path}")
    return path


def ensure_fresh_lane(lane: str) -> Path:
    directory = CAPTURES / lane
    require(not directory.exists() and not directory.is_symlink(), f"lane already exists: {directory}")
    directory.mkdir(parents=True)
    for name in (
        "started.json",
        "report.json",
        "corpus-catalog.json",
        "resource.log",
        "stdout.log",
        "stderr.log",
        "validation.log",
        "receipt.json",
    ):
        require(not (directory / name).exists(), f"lane artifact already exists: {directory / name}")
    return directory


def prepare_temp_root() -> None:
    require(not TEMP_ROOT.exists() and not TEMP_ROOT.is_symlink(),
            f"owned TMPDIR already exists: {TEMP_ROOT}")
    TEMP_ROOT.parent.mkdir(parents=True, exist_ok=True)
    TEMP_ROOT.mkdir()
    marker = TEMP_ROOT / ".litchi-0501-owned"
    marker.write_text("change-0501 baseline-capture.py\n", encoding="utf-8", newline="\n")


def cleanup_temp_root() -> None:
    if not TEMP_ROOT.exists() and not TEMP_ROOT.is_symlink():
        return
    require(not TEMP_ROOT.is_symlink(), f"owned TMPDIR became a symlink: {TEMP_ROOT}")
    marker = TEMP_ROOT / ".litchi-0501-owned"
    require(marker.is_file() and not marker.is_symlink(),
            f"owned TMPDIR marker is missing: {TEMP_ROOT}")
    shutil.rmtree(TEMP_ROOT)
    require(not TEMP_ROOT.exists() and not TEMP_ROOT.is_symlink(),
            f"owned TMPDIR cleanup left a path behind: {TEMP_ROOT}")


def workload_argv(binary: Path, lane: Path) -> list[str]:
    return [
        str(binary),
        "--workers", WORKERS,
        "--samples", str(SAMPLES),
        "--warmup", str(WARMUPS),
        "--json", rel(lane / "report.json"),
        "--corpus-manifest", rel(lane / "corpus-catalog.json"),
    ]


def timed_argv(binary: Path, lane: Path) -> list[str]:
    # Keep /usr/bin/time outside taskset, matching the historical 0465
    # protocol and retaining the complete child process resource envelope.
    return [
        "/usr/bin/time", "-v", "-o", rel(lane / "resource.log"),
        "taskset", "-c", CPU,
        *workload_argv(binary, lane),
    ]


def selected_environment() -> dict[str, str]:
    values = {
        "RUSTUP_TOOLCHAIN": "1.98.1",
        "DEBUGINFOD_URLS": "",
        "PYTHONDONTWRITEBYTECODE": "1",
        "LC_ALL": "C",
        "TMPDIR": str(TEMP_ROOT),
        "TMP": str(TEMP_ROOT),
        "TEMP": str(TEMP_ROOT),
    }
    return values


def validate_report_shape(report_path: Path, catalog_path: Path, binary: Path) -> None:
    report = read_json(report_path)
    catalog = read_json(catalog_path)
    require(report.get("schema_version") == 1, "report schema version differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), "report configuration is missing")
    require(configuration.get("samples_per_case") == SAMPLES, "report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == WARMUPS, "report warmup count differs")
    require(configuration.get("execution_workers") == [1], "report worker selection differs")
    cases = configuration.get("cases")
    require(isinstance(cases, list) and len(cases) == EXPECTED_CASES,
            f"report default case count differs: {len(cases) if isinstance(cases, list) else cases!r}")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == EXPECTED_ROWS,
            f"report result count differs: {len(results) if isinstance(results, list) else results!r}")
    keys: set[tuple[str, str]] = set()
    for index, result in enumerate(results):
        require(isinstance(result, dict), f"report result {index} is not an object")
        case = result.get("case")
        corpus = result.get("corpus")
        require(isinstance(case, str) and isinstance(corpus, dict),
                f"report result {index} has incomplete identity")
        archive_sha = corpus.get("archive_sha256")
        package_format = corpus.get("package_format")
        require(isinstance(archive_sha, str) and isinstance(package_format, str),
                f"report result {index} has incomplete corpus identity")
        key = (case, f"{package_format}:{archive_sha}")
        require(key not in keys, f"report duplicates case/corpus row: {key}")
        keys.add(key)
        elapsed = result.get("elapsed_ns")
        require(isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list),
                f"report result {index} has no elapsed samples")
        require(len(elapsed["samples"]) == SAMPLES,
                f"report result {index} sample length differs")
    require(catalog.get("manifest_version") == 2, "catalog manifest version differs")
    require(catalog.get("manifest_kind") == "corpus-catalog", "catalog kind differs")
    bindings = catalog.get("case_bindings")
    require(isinstance(bindings, list) and len(bindings) == EXPECTED_ROWS,
            "catalog binding count differs")
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict), "report omitted corpus catalog reference")
    for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
        require(reference.get(key) == catalog.get(key), f"report/catalog {key} differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), "report binary identity is missing")
    require(identity.get("binary_sha256") == sha(binary), "report binary SHA differs")
    require(identity.get("binary_bytes") == binary.stat().st_size,
            "report binary byte count differs")


def validate_with_repository_gate(lane: Path) -> tuple[list[str], int, str]:
    command = [
        sys.executable,
        "-B",
        rel(INDEX_VALIDATOR),
        "--catalog",
        rel(lane / "corpus-catalog.json"),
        "--report",
        rel(lane / "report.json"),
    ]
    result = subprocess.run(
        command,
        cwd=REPO,
        check=False,
        capture_output=True,
        text=True,
    )
    output = result.stdout + result.stderr
    (lane / "validation.log").write_text(output, encoding="utf-8")
    require(result.returncode == 0, f"coverage validator failed: {output.strip()}")
    return command, result.returncode, output


def run_static_tests(source_manifest: dict[str, Any]) -> dict[str, Any]:
    command = [sys.executable, "-B", "-m", "unittest", STATIC_TEST_MODULE]
    source_before = verify_source_manifest(source_manifest)
    result = subprocess.run(
        command,
        cwd=REPO,
        check=False,
        capture_output=True,
        text=True,
    )
    output = result.stdout + result.stderr
    path = ROOT / "coverage-tests.log"
    require(not path.exists() and not path.is_symlink(), f"artifact already exists: {path}")
    path.write_text(output, encoding="utf-8")
    require(result.returncode == 0, f"CRUD coverage tests failed: {output.strip()}")
    source_after = verify_source_manifest(source_manifest)
    require(source_after == source_before, "coverage tests changed bound source")
    return {
        "argv": command,
        "exit_code": result.returncode,
        "artifact": artifact(path),
        "source_before": source_before,
        "source_after": source_after,
    }


def run_lane(
    binary: Path,
    lane_name: str,
    source_manifest: dict[str, Any],
    observed_head: str,
) -> dict[str, Any]:
    lane = ensure_fresh_lane(lane_name)
    before_status = git_status()
    source_before = verify_source_manifest(source_manifest)
    binary_sha = sha(binary)
    binary_bytes = binary.stat().st_size
    command = timed_argv(binary, lane)
    environment = selected_environment()
    started = {
        "schema": "litchi-0501-default-baseline-start-v1",
        "lane": lane_name,
        "argv": command,
        "cwd": str(REPO),
        "binary": {"path": str(binary), "bytes": binary_bytes, "sha256": binary_sha},
        "candidate_base_revision": source_manifest["candidate_base_revision"],
        "source_manifest": source_manifest,
        "source_before": source_before,
        "observed_git_head": observed_head,
        "worktree_status_before": before_status,
        "environment": environment,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "expected_rows": EXPECTED_ROWS,
        "started_utc": now(),
    }
    write_json_exclusive(lane / "started.json", started)

    record: dict[str, Any] = {
        "schema": "litchi-0501-default-baseline-v1",
        "lane": lane_name,
        "argv": command,
        "cwd": str(REPO),
        "binary": {"path": str(binary), "bytes": binary_bytes, "sha256": binary_sha},
        "candidate_base_revision": source_manifest["candidate_base_revision"],
        "source_manifest": source_manifest,
        "source_before": source_before,
        "observed_git_head": observed_head,
        "driver_sha256": sha(Path(__file__)),
        "environment": environment,
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "expected_rows": EXPECTED_ROWS,
        "started_sha256": sha(lane / "started.json"),
        "started_utc": started["started_utc"],
        "status": "running",
    }
    output_env = os.environ.copy()
    output_env.update(environment)
    process_error: str | None = None
    try:
        prepare_temp_root()
        with (lane / "stdout.log").open("x", encoding="utf-8") as stdout, \
                (lane / "stderr.log").open("x", encoding="utf-8") as stderr:
            result = subprocess.run(
                command,
                cwd=REPO,
                env=output_env,
                stdout=stdout,
                stderr=stderr,
                check=False,
            )
        record["exit_code"] = result.returncode
        require(result.returncode == 0, f"workload exited with status {result.returncode}")
        require(sha(binary) == binary_sha, "frozen binary changed during workload")
        source_after = verify_source_manifest(source_manifest)
        require(source_after == source_before, "bound source changed during workload")
        record["source_after"] = source_after
        validate_report_shape(lane / "report.json", lane / "corpus-catalog.json", binary)
        validation_command, validation_exit, validation_output = validate_with_repository_gate(lane)
        record["validation"] = {
            "argv": validation_command,
            "exit_code": validation_exit,
            "output": validation_output,
            "report": artifact(lane / "report.json"),
            "catalog": artifact(lane / "corpus-catalog.json"),
        }
        record["status"] = "pass"
    except Exception as error:
        process_error = str(error)
        record["status"] = "failed"
    finally:
        try:
            cleanup_temp_root()
            record["temporary_directory_cleanup"] = {"path": str(TEMP_ROOT), "absent": True}
        except Exception as error:
            record["temporary_directory_cleanup"] = {"path": str(TEMP_ROOT), "absent": False}
            record["status"] = "failed"
            process_error = process_error or str(error)
        record["finished_utc"] = now()
        try:
            record["source_after"] = verify_source_manifest(source_manifest)
            record["source_unchanged"] = record["source_after"] == source_before
            if not record["source_unchanged"]:
                record["status"] = "failed"
                process_error = process_error or "bound source changed during capture"
        except Exception as error:
            record["source_unchanged"] = False
            record["source_verification_error"] = str(error)
            record["status"] = "failed"
            process_error = process_error or str(error)
        record["observed_git_head_after"] = git_revision()
        record["worktree_status_after"] = git_status()
        record["binary_after"] = artifact(binary)
        record["artifacts"] = {
            name: artifact(lane / name)
            for name in (
                "started.json",
                "report.json",
                "corpus-catalog.json",
                "resource.log",
                "stdout.log",
                "stderr.log",
                "validation.log",
            )
            if (lane / name).is_file() and not (lane / name).is_symlink()
        }
        if process_error:
            record["error"] = process_error
        write_json_exclusive(lane / "receipt.json", record)
    if record["status"] != "pass":
        raise CaptureError(f"{lane_name} failed; see {lane / 'receipt.json'}")
    return record


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("binary", help="absolute path to the frozen normal release harness")
    parser.add_argument(
        "source_manifest",
        help="source hash manifest for the final candidate (relative to the repository or absolute)",
    )
    args = parser.parse_args()
    try:
        binary = ensure_binary(args.binary)
        source_manifest_path = ensure_source_manifest(args.source_manifest)
        source_manifest = load_source_manifest(source_manifest_path)
        require(CHECKED_CATALOG.is_file() and not CHECKED_CATALOG.is_symlink(),
                f"checked catalog is missing or symlinked: {CHECKED_CATALOG}")
        require(ROOT == Path(__file__).resolve().parent, "capture root resolution changed")
        require(not (ROOT / "coverage-tests.log").exists()
                and not (ROOT / "coverage-tests.log").is_symlink(),
                f"artifact already exists: {ROOT / 'coverage-tests.log'}")
        require(not (ROOT / "baseline-capture-summary.json").exists()
                and not (ROOT / "baseline-capture-summary.json").is_symlink(),
                f"artifact already exists: {ROOT / 'baseline-capture-summary.json'}")
        observed_head = git_revision()
        verify_source_manifest(source_manifest)
        records = []
        for lane in LANES:
            records.append(run_lane(binary, lane, source_manifest, observed_head))
        static_tests = run_static_tests(source_manifest)
        summary = {
            "schema": "litchi-0501-default-baseline-summary-v1",
            "driver": artifact(Path(__file__)),
            "binary": artifact(binary),
            "candidate_base_revision": source_manifest["candidate_base_revision"],
            "source_manifest": source_manifest,
            "observed_git_head": observed_head,
            "lanes": records,
            "static_coverage_tests": static_tests,
            "temporary_directory": {"path": str(TEMP_ROOT), "absent": not TEMP_ROOT.exists()},
            "finished_utc": now(),
        }
        write_json_exclusive(ROOT / "baseline-capture-summary.json", summary)
        print(json.dumps({"status": "pass", "lanes": list(LANES)}))
        return 0
    except Exception as error:
        print(f"baseline capture failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
