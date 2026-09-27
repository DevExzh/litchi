"""Pure offline custody and replay for the retained 0781 CI evidence.

The CI capture and quality drivers are deliberately root-only programs.  This
module is their read-only counterpart: :func:`analyze` checks the retained
receipts, replays the validators' in-memory APIs, and returns a stable summary.
It never starts a process and never writes an output file.  Absolute paths in
old receipts are resolved through ``origin.json`` so the packet remains
auditable after the owned worktree and build target have been removed.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import re
from pathlib import Path
from types import ModuleType
from typing import Any


PACKET = Path(__file__).resolve().parent
BASELINE_BUILD = "ci-baseline-build"
CAPTURE_DIR = "ci-capture"
QUALITY_JSON = "ci-quality.json"
SMOKE_ROWS = 41
FULL_ROWS = 213
QUALITY_GATES = 6
SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")


class CIValidationError(RuntimeError):
    """Raised when retained CI evidence is missing or contradictory."""


def fail(message: str) -> None:
    raise CIValidationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid {label} {path}: {error}")


def _object(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} must be an object")
    return value


def _list(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} must be a list")
    return value


def _string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} must be a non-empty string")
    return value


def _integer(value: Any, label: str, *, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
            f"{label} must be an integer >= {minimum}")
    return value


def _boolean(value: Any, label: str) -> bool:
    require(isinstance(value, bool), f"{label} must be a boolean")
    return value


def _sha(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256.fullmatch(value) is not None,
            f"{label} must be a lowercase SHA-256")
    return value


def _revision(value: Any, label: str) -> str:
    require(isinstance(value, str) and REVISION.fullmatch(value) is not None,
            f"{label} must be a 40-character commit revision")
    return value


def _finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} must be finite")
    return float(value)


def _digest(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                value.update(chunk)
    except OSError as error:
        fail(f"cannot hash retained artifact {path}: {error}")
    return value.hexdigest()


def _identity(value: dict[str, Any], label: str) -> dict[str, Any]:
    # Empty capture logs are valid retained artifacts; binary custody applies
    # its own positive-size check below through the captured identity.
    _integer(value.get("bytes"), f"{label}.bytes")
    digest = _sha(value.get("sha256"), f"{label}.sha256")
    return {"bytes": value["bytes"], "sha256": digest}


class Context:
    """Packet and relocation paths used by all custody checks."""

    def __init__(self, packet: Path):
        packet = packet.resolve()
        require(packet.is_dir(), f"CI packet is not a directory: {packet}")
        self.packet = packet
        # change-0781 is docs/performance/results/<packet>, so four parents
        # above the packet is the checkout root.
        require(len(packet.parents) >= 4, f"CI packet path is too shallow: {packet}")
        self.root = packet.parents[3].resolve()
        origin = _object(_read_json(packet / "origin.json", "origin.json"), "origin.json")
        owned = _string(origin.get("owned_worktree"), "origin owned_worktree")
        target = _string(origin.get("target"), "origin target")
        base = _string(origin.get("base"), "origin base")
        _revision(base, "origin base")
        self.origin = origin
        self.owned = Path(owned).resolve()
        self.target = Path(target).resolve()
        self.base = base

    def normalize(self, value: str) -> str:
        """Map only the archived owned-worktree prefix to this checkout."""

        old = str(self.owned)
        current = str(self.root)
        if value == old:
            return current
        if value.startswith(old + "/"):
            return current + value[len(old):]
        return value

    def candidates(self, value: str) -> list[Path]:
        raw = Path(value)
        candidates: list[Path] = []
        if raw.is_absolute():
            mapped = Path(self.normalize(value))
            if mapped != raw:
                candidates.append(mapped)
            candidates.append(raw)
            try:
                relative = raw.relative_to(self.packet)
            except ValueError:
                relative = None
            if relative is not None:
                candidates.append(self.packet / relative)
        else:
            text = value.replace("\\", "/")
            prefix = "docs/performance/results/change-0781/"
            if text.startswith(prefix):
                candidates.append(self.packet / text[len(prefix):])
            candidates.extend((self.packet / raw, self.root / raw))
        unique: list[Path] = []
        for candidate in candidates:
            candidate = candidate.absolute()
            if candidate not in unique:
                unique.append(candidate)
        return unique

    def resolve(self, value: str, *, packet_bound: bool) -> Path | None:
        for candidate in self.candidates(value):
            if candidate.is_symlink():
                continue
            if not candidate.is_file():
                continue
            resolved = candidate.resolve()
            if packet_bound:
                try:
                    resolved.relative_to(self.packet)
                except ValueError:
                    continue
            return resolved
        return None

    def relative(self, path: Path) -> str:
        try:
            return str(path.resolve().relative_to(self.packet))
        except ValueError:
            return str(path)

    def artifact(
        self,
        value: Any,
        label: str,
        *,
        packet_bound: bool = True,
        allow_missing: bool = False,
    ) -> tuple[Path | None, dict[str, Any]]:
        receipt = _object(value, label)
        raw_path = _string(receipt.get("path"), f"{label}.path")
        expected = _identity(receipt, label)
        path = self.resolve(raw_path, packet_bound=packet_bound)
        if path is None:
            if allow_missing:
                return None, expected
            fail(f"missing {label}: {raw_path}")
        require(path.stat().st_size == expected["bytes"], f"{label}.bytes changed")
        require(_digest(path) == expected["sha256"], f"{label}.sha256 changed")
        return path, expected


def _cleanup_verified(cleanup: Any) -> bool:
    return isinstance(cleanup, dict) and any(
        cleanup.get(key) is True
        for key in ("verified", "executables_verified_before_removal",
                    "binary_removal_verified")
    )


def _cleanup_contains(value: Any, receipt: dict[str, Any]) -> bool:
    wanted = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
              receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(wanted[0], str) and isinstance(wanted[1], int)
            and not isinstance(wanted[1], bool) and SHA256.fullmatch(wanted[2] or "")):
        return False
    if isinstance(value, dict):
        actual = (value.get("path"), value.get("bytes", value.get("size")),
                  value.get("sha256", value.get("digest")))
        if actual == wanted:
            return True
        return any(_cleanup_contains(child, receipt) for child in value.values())
    if isinstance(value, list):
        return any(_cleanup_contains(child, receipt) for child in value)
    return False


def _load_cleanup(ctx: Context) -> tuple[Any, bool]:
    path = ctx.packet / "cleanup.json"
    if not path.is_file() or path.is_symlink():
        return None, False
    cleanup = _object(_read_json(path, "cleanup.json"), "cleanup.json")
    return cleanup, _cleanup_verified(cleanup)


def _binary_custody(
    ctx: Context,
    receipt: dict[str, Any],
    label: str,
    cleanup: Any,
    cleanup_verified: bool,
) -> dict[str, Any]:
    _integer(receipt.get("bytes"), f"{label}.bytes", minimum=1)
    path, identity = ctx.artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is not None:
        return identity
    require(cleanup_verified, f"{label} is missing without cleanup verification")
    require(_cleanup_contains(cleanup, receipt), f"{label} lacks an exact cleanup witness")
    # Do not expose whether the binary was retained or removed: relocation and
    # cleanup must not change the deterministic analysis result.
    return identity


def _source_manifest(value: Any, label: str) -> dict[str, Any]:
    source = _object(value, label)
    revision = _string(source.get("revision"), f"{label}.revision")
    _revision(revision, f"{label}.revision")
    files = _object(source.get("files"), f"{label}.files")
    require(files, f"{label}.files is empty")
    for name, digest in files.items():
        _string(name, f"{label} file name")
        _sha(digest, f"{label}.{name}")
    return {"revision": revision, "files": dict(files)}


def _artifact_rel(ctx: Context, path: Path) -> str:
    return ctx.relative(path)


def _module(path: Path, name: str) -> ModuleType:
    """Load a trusted repository validator without executing its CLI."""

    require(path.is_file() and not path.is_symlink(), f"missing validator source: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load validator: {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:  # validator errors are reported with packet context
        fail(f"cannot import validator {path}: {error}")
    return module


def _replay_matrix(ctx: Context, report_path: Path, *, mode: str, samples: int) -> dict[str, Any]:
    matrix_path = ctx.root / "tools/validate_perf_default_matrix.py"
    module = _module(matrix_path, "litchi_ci_matrix_validator")
    require(hasattr(module, "load_manifest") and hasattr(module, "validate_report"),
            "default matrix validator API changed")
    manifest_path = ctx.root / "docs/performance/results/perf-regression-default-manifest-v1.json"
    try:
        manifest = module.load_manifest(manifest_path)
        summary = module.validate_report(
            report_path,
            manifest,
            mode=mode,
            samples=samples,
            shape="tiny" if mode == "smoke" else None,
            payload="compressible" if mode == "smoke" else None,
        )
    except Exception as error:
        fail(f"{mode} matrix replay failed: {error}")
    summary = _object(summary, f"{mode} matrix summary")
    result_count = _integer(summary.get("result_count"), f"{mode} result_count", minimum=1)
    case_count = _integer(summary.get("case_count"), f"{mode} case_count", minimum=1)
    expected = SMOKE_ROWS if mode == "smoke" else FULL_ROWS
    require(result_count == expected,
            f"{mode} matrix has {result_count} rows; expected {expected}")
    require(case_count == SMOKE_ROWS,
            f"{mode} matrix has {case_count} cases; expected {SMOKE_ROWS}")
    digest = _sha(summary.get("result_keys_sha256"), f"{mode} result key digest")
    return {"cases": case_count, "rows": result_count, "result_keys_sha256": digest}


def _validate_baseline(
    ctx: Context,
    cleanup: Any,
    cleanup_verified: bool,
) -> dict[str, Any]:
    directory = ctx.packet / BASELINE_BUILD
    build = _object(_read_json(directory / "build.json", "CI baseline build.json"),
                    "CI baseline build.json")
    expected_command = [
        "cargo", "build", "--release", "--locked", "--offline",
        "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin",
        "litchi-perf-baseline",
    ]
    require(build.get("command") == expected_command, "CI baseline build command changed")
    require(build.get("exit_code") == 0, "CI baseline harness build failed")
    environment = _object(build.get("environment"), "CI baseline build environment")
    require(environment.get("CARGO_BUILD_JOBS") == "2",
            "CI baseline CARGO_BUILD_JOBS changed")
    require(environment.get("CARGO_INCREMENTAL") == "0",
            "CI baseline CARGO_INCREMENTAL changed")
    target_dir = str(ctx.target / "harness")
    require(environment.get("CARGO_TARGET_DIR") == target_dir,
            "CI baseline target directory changed")
    _finite(build.get("started"), "CI baseline build started")
    _finite(build.get("ended"), "CI baseline build ended")
    require(build["ended"] >= build["started"], "CI baseline build time is reversed")
    _, build_log = ctx.artifact(build.get("log"), "CI baseline build log")
    source_path, source_receipt = ctx.artifact(
        build.get("source"), "CI baseline source artifact"
    )
    require(source_path is not None, "CI baseline source artifact is missing")
    source = _source_manifest(_read_json(source_path, "CI baseline source"),
                              "CI baseline source")
    require(source["revision"] == ctx.base,
            "CI baseline source is not the origin base revision")

    harness = _object(_read_json(directory / "harness-source.json", "CI harness source"),
                      "CI harness source")
    require(harness.get("revision") == ctx.base,
            "CI harness source revision is not the origin base")
    require(harness.get("working_tree_matches_base") is True,
            "CI harness source does not attest base-code custody")
    harness_files = _object(harness.get("files"), "CI harness source files")
    require(harness_files, "CI harness source files are empty")
    for name, expected in harness_files.items():
        _string(name, "CI harness source file name")
        _sha(expected, f"CI harness source {name}")
        path = ctx.root / name
        require(path.is_file() and not path.is_symlink(),
                f"missing CI harness source file: {name}")
        require(_digest(path) == expected, f"CI harness source changed: {name}")
    lock_path, lock = ctx.artifact(harness.get("lock"), "CI harness Cargo.lock",
                                   packet_bound=False)
    manifest_path, manifest = ctx.artifact(harness.get("manifest"),
                                            "CI harness Cargo.toml", packet_bound=False)
    require(lock_path is not None and manifest_path is not None,
            "CI harness lock or manifest is missing")
    require(harness_files.get("tools/perf-baseline/Cargo.lock") == lock["sha256"],
            "CI harness lock receipt differs from source census")
    require(harness_files.get("tools/perf-baseline/Cargo.toml") == manifest["sha256"],
            "CI harness manifest receipt differs from source census")

    binary = _object(_read_json(directory / "binary.json", "CI baseline binary"),
                     "CI baseline binary")
    binary_path = _string(binary.get("path"), "CI baseline binary.path")
    expected_binary_path = str(ctx.target / "harness/release/litchi-perf-baseline")
    require(binary_path == expected_binary_path, "CI baseline binary path changed")
    binary_identity = _binary_custody(ctx, binary, "CI baseline binary", cleanup,
                                      cleanup_verified)
    return {
        "source": source,
        "source_artifact": source_receipt,
        "harness_source_files": len(harness_files),
        "binary": binary_identity,
        "build_log": build_log,
        "build_scope": "base",
    }


def _report_identity(
    ctx: Context,
    report: dict[str, Any],
    binary: dict[str, Any],
    mode: str,
    samples: int,
) -> None:
    require(report.get("schema_version") == 1, f"{mode} report schema changed")
    tool = _object(report.get("tool"), f"{mode} report tool")
    require(tool.get("name") == "litchi-perf-baseline", f"{mode} report tool changed")
    require(tool.get("binary") == "litchi-perf-baseline", f"{mode} report binary changed")
    require(tool.get("profile") == "release", f"{mode} report profile changed")
    identity = _object(report.get("binary_identity"), f"{mode} binary identity")
    require(identity.get("path") == binary["path"], f"{mode} binary path identity changed")
    require(identity.get("binary_sha256") == binary["sha256"],
            f"{mode} binary SHA identity changed")
    require(identity.get("binary_bytes") == binary["bytes"],
            f"{mode} binary byte identity changed")
    require(identity.get("executable") is True and identity.get("profile") == "release",
            f"{mode} binary executable identity changed")
    environment = _object(report.get("environment"), f"{mode} report environment")
    require(environment.get("git_revision") == ctx.base,
            f"{mode} report git revision is not the base")
    configuration = _object(report.get("configuration"), f"{mode} report configuration")
    require(configuration.get("samples_per_case") == samples,
            f"{mode} report samples_per_case changed")


def _expected_capture_command(
    ctx: Context,
    binary: dict[str, Any],
    mode: str,
    report: Path,
    catalog: Path | None,
) -> list[str]:
    binary_path = str(ctx.target / "harness/release/litchi-perf-baseline")
    output = str(report)
    if mode == "smoke":
        return [binary_path, "--samples", "2", "--shape", "tiny", "--payload",
                "compressible", "--writer-shape", "tiny", "--xlsx-shape", "tiny",
                "--semantic-shape", "tiny", "--json", output]
    require(catalog is not None, "full capture catalog is missing")
    return [binary_path, "--samples", "15", "--corpus-manifest", str(catalog),
            "--json", output]


def _validate_capture(
    ctx: Context,
    baseline: dict[str, Any],
    cleanup: Any,
    cleanup_verified: bool,
) -> dict[str, Any]:
    directory = ctx.packet / CAPTURE_DIR
    complete = _object(_read_json(directory / "complete.json", "CI capture complete.json"),
                       "CI capture complete.json")
    require(complete.get("processes") == 2, "CI capture process count changed")
    receipt_path, receipt_identity = ctx.artifact(complete.get("receipts"),
                                                  "CI capture receipts")
    require(receipt_path is not None, "CI capture receipts are missing")
    rows = _list(_read_json(receipt_path, "CI capture receipts"), "CI capture receipts")
    require(len(rows) == 2, "CI capture receipt count changed")
    source = _source_manifest(_read_json(directory / "source.json", "CI capture source"),
                             "CI capture source")
    require(source == baseline["source"],
            "CI capture source does not match the base CI build source")
    capture_source = {
        "bytes": (directory / "source.json").stat().st_size,
        "sha256": _digest(directory / "source.json"),
    }
    require(capture_source == baseline["source_artifact"],
            "CI capture source artifact differs from the base CI source receipt")
    modes = ["smoke", "full"]
    summaries: dict[str, Any] = {}
    full_catalog_path: Path | None = None
    for index, (row_value, mode) in enumerate(zip(rows, modes)):
        row = _object(row_value, f"CI capture receipt {mode}")
        require(row.get("mode") == mode, f"CI capture receipt {index} mode changed")
        require(row.get("exit_code") == 0, f"CI {mode} capture failed")
        _finite(row.get("started"), f"CI {mode} capture started")
        _finite(row.get("ended"), f"CI {mode} capture ended")
        require(row["ended"] >= row["started"], f"CI {mode} capture time is reversed")
        binary = _object(row.get("binary"), f"CI {mode} binary receipt")
        require(binary == {
            "bytes": baseline["binary"]["bytes"],
            "path": str(ctx.target / "harness/release/litchi-perf-baseline"),
            "sha256": baseline["binary"]["sha256"],
        }, f"CI {mode} binary receipt differs from the baseline binary")
        _binary_custody(ctx, binary, f"CI {mode} binary", cleanup, cleanup_verified)
        log_path, log = ctx.artifact(row.get("log"), f"CI {mode} capture log")
        report_path, report_identity = ctx.artifact(row.get("report"),
                                                   f"CI {mode} report")
        require(report_path is not None and log_path is not None,
                f"CI {mode} capture artifact is missing")
        catalog_path = None
        catalog_identity = None
        if mode == "full":
            catalog_path, catalog_identity = ctx.artifact(row.get("catalog"),
                                                          "CI full corpus catalog")
            require(catalog_path is not None and catalog_identity is not None,
                    "CI full corpus catalog is missing")
            full_catalog_path = catalog_path
        else:
            require("catalog" not in row, "CI smoke capture unexpectedly has a catalog")

        require(ctx.relative(report_path) == f"ci-capture/{mode}.json",
                f"CI {mode} report path changed")
        if mode == "full":
            require(ctx.relative(catalog_path) == "ci-capture/full.corpus-manifest-v2.json",
                    "CI full corpus catalog path changed")

        report = _object(_read_json(report_path, f"CI {mode} report"),
                         f"CI {mode} report")
        _report_identity(ctx, report, binary, mode, 2 if mode == "smoke" else 15)
        expected = _expected_capture_command(ctx, binary, mode, report_path, catalog_path)
        actual_command = row.get("command")
        require(isinstance(actual_command, list)
                and all(isinstance(item, str) for item in actual_command),
                f"CI {mode} command is malformed")
        require([ctx.normalize(item) for item in actual_command]
                == [ctx.normalize(item) for item in expected],
                f"CI {mode} command changed")
        matrix = _replay_matrix(ctx, report_path, mode=mode,
                                samples=2 if mode == "smoke" else 15)
        summaries[mode] = {
            "rows": matrix["rows"],
            "cases": matrix["cases"],
            "result_keys_sha256": matrix["result_keys_sha256"],
            "report": report_identity,
            "log": log,
            **({"catalog": catalog_identity} if catalog_identity else {}),
        }

    require(full_catalog_path is not None, "CI full catalog was not retained")
    binding_module = _module(ctx.root / "tools/validate_perf_corpus_binding.py",
                             "litchi_ci_corpus_validator")
    require(hasattr(binding_module, "validate_paths"),
            "corpus binding validator API changed")
    try:
        corpus_count, binding_count = binding_module.validate_paths(
            full_catalog_path.parent / "full.json", full_catalog_path
        )
    except Exception as error:
        fail(f"full corpus binding replay failed: {error}")
    require((corpus_count, binding_count) == (43, FULL_ROWS),
            "full corpus binding counts changed")
    summaries["full"]["corpus_binding"] = {
        "corpora": corpus_count,
        "bindings": binding_count,
    }
    return {
        "source": {"revision": source["revision"], "files": len(source["files"])},
        "source_artifact": capture_source,
        "receipts": receipt_identity,
        "processes": len(rows),
        "modes": summaries,
    }


def _expected_quality_commands(ctx: Context, quality_dir: Path) -> list[list[str]]:
    capture = ctx.packet / CAPTURE_DIR
    return [
        ["python3", "-B", "-m", "unittest", "tools.test_validate_perf_default_matrix",
         "tools.test_perf_workflow_policy", "tools.test_crud_coverage_index"],
        ["python3", "-B", "tools/validate_perf_default_matrix.py", "--manifest",
         "docs/performance/results/perf-regression-default-manifest-v1.json", "--report",
         str(capture / "smoke.json"), "--mode", "smoke", "--samples", "2", "--shape",
         "tiny", "--payload", "compressible"],
        ["python3", "-B", "tools/validate_perf_default_matrix.py", "--manifest",
         "docs/performance/results/perf-regression-default-manifest-v1.json", "--report",
         str(capture / "full.json"), "--mode", "full", "--samples", "15"],
        ["python3", "-B", "tools/validate_crud_coverage_index.py", "--index",
         "docs/performance/crud-coverage-index-v2.json"],
        ["python3", "-B", "tools/validate_perf_corpus_binding.py", "--report",
         str(capture / "full.json"), "--catalog", str(capture / "full.corpus-manifest-v2.json")],
        ["python3", "-B", "tools/validate_crud_coverage_index.py", "--index",
         "docs/performance/crud-coverage-index-v2.json", "--catalog",
         str(capture / "full.corpus-manifest-v2.json"), "--selector-source",
         "tools/perf-baseline/src/lib.rs", "--checklist", "docs/CRUD_Scenario_Checklist.md",
         "--repo-root", ".", "--report", str(capture / "full.json")],
    ]


def _log_text(path: Path, label: str) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {label}: {error}")


def _validate_quality(ctx: Context, capture: dict[str, Any]) -> dict[str, Any]:
    quality = _object(_read_json(ctx.packet / QUALITY_JSON, QUALITY_JSON), QUALITY_JSON)
    rows = _list(quality.get("rows"), "CI quality rows")
    require(len(rows) == QUALITY_GATES,
            f"CI quality has {len(rows)} gates; expected {QUALITY_GATES}")
    source_path, source_receipt = ctx.artifact(quality.get("source"),
                                               "CI quality source artifact")
    require(source_path is not None, "CI quality source artifact is missing")
    quality_dir = source_path.parent
    require(quality_dir.name.startswith("ci-quality-"),
            "CI quality source is outside a quality attempt directory")
    source = _object(_read_json(source_path, "CI quality source"), "CI quality source")
    require(source, "CI quality source is empty")
    for name, expected in source.items():
        _string(name, "CI quality source file name")
        _sha(expected, f"CI quality source {name}")
        path = ctx.root / name
        require(path.is_file() and not path.is_symlink(),
                f"missing CI quality source file: {name}")
        require(_digest(path) == expected, f"CI quality source changed: {name}")

    checks_path = quality_dir / "checks.json"
    checks = _list(_read_json(checks_path, "CI quality checks"), "CI quality checks")
    require(checks == rows, "CI quality checks.json differs from ci-quality.json")
    expected_commands = _expected_quality_commands(ctx, quality_dir)
    logs: list[dict[str, Any]] = []
    test_count: int | None = None
    matrix_summaries: dict[str, dict[str, Any]] = {}
    for index, (row_value, expected_command) in enumerate(zip(rows, expected_commands)):
        row = _object(row_value, f"CI quality gate {index + 1}")
        require(row.get("exit_code") == 0, f"CI quality gate {index + 1} failed")
        _finite(row.get("started"), f"CI quality gate {index + 1} started")
        _finite(row.get("ended"), f"CI quality gate {index + 1} ended")
        require(row["ended"] >= row["started"],
                f"CI quality gate {index + 1} time is reversed")
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(item, str) for item in command),
                f"CI quality gate {index + 1} command is malformed")
        require([ctx.normalize(item) for item in command]
                == [ctx.normalize(item) for item in expected_command],
                f"CI quality gate {index + 1} command changed")
        log_path, log_identity = ctx.artifact(row.get("log"),
                                               f"CI quality gate {index + 1} log")
        require(log_path is not None, f"CI quality gate {index + 1} log is missing")
        require(log_path.parent.resolve() == quality_dir.resolve(),
                f"CI quality gate {index + 1} log escaped quality directory")
        text = _log_text(log_path, f"CI quality gate {index + 1} log")
        if index == 0:
            match = re.search(
                r"Ran ([0-9]+) tests in [0-9]+(?:\.[0-9]+)?s\n\nOK\n?\Z",
                text,
            )
            require(match is not None, "CI unit-test log is not a passing unittest result")
            test_count = int(match.group(1))
            require(test_count > 0, "CI unit-test log reported no tests")
        elif index in (1, 2):
            mode = "smoke" if index == 1 else "full"
            expected_rows = SMOKE_ROWS if mode == "smoke" else FULL_ROWS
            match = re.fullmatch(
                rf"validated default performance matrix: {SMOKE_ROWS} cases, "
                rf"{expected_rows} rows, ([0-9a-f]{{64}})\n",
                text,
            )
            require(match is not None, f"CI {mode} matrix log is not a passing result")
            summary = capture["modes"][mode]
            require(match.group(1) == summary["result_keys_sha256"],
                    f"CI {mode} matrix log digest differs from replay")
            matrix_summaries[mode] = {"rows": expected_rows, "cases": SMOKE_ROWS,
                                      "result_keys_sha256": match.group(1)}
        elif index == 3:
            require(text == "validated non-iWork CRUD coverage index: 15 categories, "
                    "33 mapped selectors (contract-only; no run timing report supplied)\n",
                    "CI contract-only CRUD log changed")
        elif index == 4:
            require(text == "validated schema-2 corpus catalog: 43 corpora, 213 bindings\n",
                    "CI corpus binding log changed")
        elif index == 5:
            require(text == "validated non-iWork CRUD coverage index: 15 categories, "
                    "33 mapped selectors and bound full-run timing report\n",
                    "CI full CRUD log changed")
        logs.append(log_identity)
    require(test_count is not None, "CI unit-test count was not recorded")
    crud_module = _module(ctx.root / "tools/validate_crud_coverage_index.py",
                          "litchi_ci_crud_validator")
    require(hasattr(crud_module, "validate_paths"), "CRUD validator API changed")
    index_path = ctx.root / "docs/performance/crud-coverage-index-v2.json"
    catalog_path = ctx.root / "docs/performance/results/perf-corpus-manifest-v2.json"
    selector_path = ctx.root / "tools/perf-baseline/src/lib.rs"
    checklist_path = ctx.root / "docs/CRUD_Scenario_Checklist.md"
    try:
        contract_counts = crud_module.validate_paths(
            index_path, catalog_path, selector_path, checklist_path, repo_root=ctx.root
        )
        full_counts = crud_module.validate_paths(
            index_path,
            ctx.packet / CAPTURE_DIR / "full.corpus-manifest-v2.json",
            selector_path,
            checklist_path,
            repo_root=ctx.root,
            report_path=ctx.packet / CAPTURE_DIR / "full.json",
        )
    except Exception as error:
        fail(f"CRUD coverage replay failed: {error}")
    require(contract_counts == (15, 33), "contract-only CRUD counts changed")
    require(full_counts == (15, 33), "full-run CRUD counts changed")
    return {
        "gates": len(rows),
        "tests": test_count,
        "source_files": len(source),
        "source": source_receipt,
        "checks": {"bytes": checks_path.stat().st_size, "sha256": _digest(checks_path)},
        "logs": logs,
        "matrix": matrix_summaries,
        "crud": {"contract": {"categories": contract_counts[0],
                                "selectors": contract_counts[1]},
                 "full": {"categories": full_counts[0], "selectors": full_counts[1]}},
    }


def analyze(packet_root: Path | str = PACKET) -> dict[str, Any]:
    """Replay retained CI evidence and return a deterministic stable summary."""

    ctx = Context(Path(packet_root))
    cleanup, cleanup_verified = _load_cleanup(ctx)
    baseline = _validate_baseline(ctx, cleanup, cleanup_verified)
    capture = _validate_capture(ctx, baseline, cleanup, cleanup_verified)
    quality = _validate_quality(ctx, capture)
    return {
        "schema": "litchi.performance.0781.ci-custody.v1",
        "baseline": {
            "revision": ctx.base,
            "scope": baseline["build_scope"],
            "source_files": len(baseline["source"]["files"]),
            "harness_source_files": baseline["harness_source_files"],
            "binary": baseline["binary"],
            "build_log": baseline["build_log"],
        },
        "capture": capture,
        "quality": quality,
        "cleanup_witness_required": True,
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("packet", nargs="?", type=Path, default=PACKET)
    arguments = parser.parse_args()
    try:
        print(json.dumps(analyze(arguments.packet), indent=2, sort_keys=True))
    except CIValidationError as error:
        parser.exit(1, f"CI custody validation failed: {error}\n")
