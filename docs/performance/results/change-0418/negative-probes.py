#!/usr/bin/env python3
"""Replay an isolated 0418 export and require concrete corruptions to fail.

The formal capture is intentionally copied into a temporary, portable bundle.
The copy has no build binaries or source worktrees: those paths are replaced
with absolute placeholders which the evidence verifier is required to accept
as opaque provenance.  A normal and an allocator report are retained through
the verifier's zstd sidecar path when zstd is available (as it is for the
published capture).  Each mutation is applied to a fresh copy and its
artifact digest is repaired when necessary, so a rejection exercises the
semantic check rather than only the outer file hash.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable


SCRIPT_ROOT = Path(__file__).resolve().parent
JSON_INDENT = 2
REQUIRED_TOOL_FILES = (
    "perf_abba_summary.py",
    "perf_compare.py",
    "perf_resource_profile.py",
    "validate_perf_corpus_binding.py",
)
REPORT_ARTIFACT_KEYS = ("report", "catalog", "time_v", "stdout", "stderr")
REPORT_PHASES = ("preflight", "normal", "allocator")


def sha256(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def json_bytes(value: Any) -> bytes:
    return (
        json.dumps(
            value,
            indent=JSON_INDENT,
            sort_keys=True,
            ensure_ascii=False,
            allow_nan=False,
        )
        + "\n"
    ).encode("utf-8")


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value: Any) -> None:
    path.write_bytes(json_bytes(value))


def artifact_storage(bundle: Path, relative: str) -> tuple[Path, bool]:
    """Return the stored path and whether it is a zstd sidecar."""

    raw = bundle / relative
    if raw.is_file():
        return raw, False
    for suffix in (".zst", ".zstd"):
        sidecar = Path(str(raw) + suffix)
        if sidecar.is_file():
            return sidecar, True
    raise RuntimeError(f"missing artifact in export: {relative}")


def read_artifact(bundle: Path, relative: str) -> bytes:
    path, compressed = artifact_storage(bundle, relative)
    if not compressed:
        return path.read_bytes()
    process = subprocess.run(
        ["zstd", "-q", "-d", "-c", str(path)],
        capture_output=True,
        check=False,
        timeout=120,
    )
    if process.returncode != 0:
        detail = process.stderr.decode("utf-8", "replace").strip()
        raise RuntimeError(f"cannot decode {relative}.zst: {detail}")
    return process.stdout


def write_artifact(bundle: Path, relative: str, value: bytes) -> None:
    path, compressed = artifact_storage(bundle, relative)
    if not compressed:
        path.write_bytes(value)
        return
    process = subprocess.run(
        ["zstd", "-q", "-c"],
        input=value,
        capture_output=True,
        check=False,
        timeout=120,
    )
    if process.returncode != 0:
        detail = process.stderr.decode("utf-8", "replace").strip()
        raise RuntimeError(f"cannot encode {relative}.zst: {detail}")
    path.write_bytes(process.stdout)


def replace_strings(value: Any, replacements: dict[str, str]) -> Any:
    if isinstance(value, dict):
        return {key: replace_strings(child, replacements) for key, child in value.items()}
    if isinstance(value, list):
        return [replace_strings(child, replacements) for child in value]
    if isinstance(value, str):
        for old, new in sorted(replacements.items(), key=lambda item: -len(item[0])):
            value = value.replace(old, new)
    return value


def copy_source_bundle(source: Path, temporary: Path, repo_root: Path) -> Path:
    """Copy only verifier inputs and the standard-library verifier modules."""

    temporary.mkdir(parents=True, exist_ok=True)
    bundle = temporary / "bundle"
    bundle.mkdir()
    for name in (
        "protocol.json",
        "roles.json",
        "build-identity.json",
        "capture.json",
        "summary.json",
        "allocation-metrics.json",
    ):
        path = source / name
        if path.is_file():
            shutil.copy2(path, bundle / name)
    for directory in REPORT_PHASES:
        path = source / directory
        if path.is_dir():
            shutil.copytree(path, bundle / directory)

    shutil.copy2(source / "verify.py", bundle / "verify.py")
    tools = temporary / "tools"
    tools.mkdir()
    for name in REQUIRED_TOOL_FILES:
        shutil.copy2(repo_root / "tools" / name, tools / name)
    return bundle


def find_repo_root() -> Path:
    for candidate in (SCRIPT_ROOT, *SCRIPT_ROOT.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    raise RuntimeError("cannot locate the repository tools directory")


def formal_report(phase: str, selector: str) -> str:
    return f"{phase}/a1-{selector}.json"


def capture_for_artifact(bundle: Path, relative: str) -> tuple[dict[str, Any], str]:
    capture_path = bundle / "capture.json"
    capture = read_json(capture_path)
    for run in capture["runs"]:
        for key in REPORT_ARTIFACT_KEYS:
            if run[key] == relative:
                return capture, key
    raise RuntimeError(f"capture has no run for artifact {relative}")


def repair_artifact_hash(bundle: Path, relative: str, data: bytes) -> None:
    capture_path = bundle / "capture.json"
    capture, key = capture_for_artifact(bundle, relative)
    for run in capture["runs"]:
        if run[key] == relative:
            run["artifact_sha256"][key] = sha256(data)
    write_json(capture_path, capture)


def mutate_json_artifact(
    bundle: Path, relative: str, mutation: Callable[[dict[str, Any]], None]
) -> None:
    data = read_artifact(bundle, relative)
    value = json.loads(data)
    if not isinstance(value, dict):
        raise RuntimeError(f"JSON artifact is not an object: {relative}")
    mutation(value)
    data = json_bytes(value)
    write_artifact(bundle, relative, data)
    repair_artifact_hash(bundle, relative, data)


def mutate_text_artifact(
    bundle: Path, relative: str, mutation: Callable[[str], str]
) -> None:
    data = read_artifact(bundle, relative)
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        raise RuntimeError(f"text artifact is not UTF-8: {relative}") from error
    data = mutation(text).encode("utf-8")
    write_artifact(bundle, relative, data)
    repair_artifact_hash(bundle, relative, data)


def mutate_capture(bundle: Path, mutation: Callable[[dict[str, Any]], None]) -> None:
    path = bundle / "capture.json"
    value = read_json(path)
    mutation(value)
    write_json(path, value)


def portableize(bundle: Path, source: Path) -> None:
    """Remove dependence on the source checkout and captured binaries."""

    build_path = bundle / "build-identity.json"
    roles_path = bundle / "roles.json"
    capture_path = bundle / "capture.json"
    build = read_json(build_path)
    roles = read_json(roles_path)
    capture = read_json(capture_path)
    source_root = source.resolve().as_posix()
    portable_root = "/__litchi_0418_portable_export__"
    replacements: dict[str, str] = {source_root: portable_root}
    for role in ("control", "candidate"):
        old_worktree = build["roles"][role]["source"]["worktree"]
        replacements[old_worktree] = f"{portable_root}/sources/{role}"
        for phase in ("normal", "allocator"):
            old_binary = build["roles"][role]["binaries"][phase]["path"]
            replacements[old_binary] = (
                f"{portable_root}/binaries/{role}-{phase}"
            )

    write_json(roles_path, replace_strings(roles, replacements))
    write_json(build_path, replace_strings(build, replacements))

    # Reports bind their binary path, while time -v binds the binary and the
    # report/catalog output paths in its Command line.  Rewrite both before
    # repairing the run-level artifact hashes.
    for run in capture["runs"]:
        for key in ("report", "catalog"):
            relative = run[key]
            try:
                data = read_artifact(bundle, relative)
            except RuntimeError:
                continue
            if key == "report":
                value = json.loads(data)
                data = json_bytes(replace_strings(value, replacements))
                write_artifact(bundle, relative, data)
            repair_artifact_hash(bundle, relative, data)
        relative = run["time_v"]
        data = read_artifact(bundle, relative)
        text = data.decode("utf-8")
        for old, new in sorted(replacements.items(), key=lambda item: -len(item[0])):
            text = text.replace(old, new)
        data = text.encode("utf-8")
        write_artifact(bundle, relative, data)
        repair_artifact_hash(bundle, relative, data)

    capture = replace_strings(read_json(capture_path), replacements)
    build_digest = sha256(build_path.read_bytes())
    capture["build_identity"]["sha256"] = build_digest
    for run in capture["runs"]:
        for key in REPORT_ARTIFACT_KEYS:
            relative = run[key]
            run["artifact_sha256"][key] = sha256(read_artifact(bundle, relative))
    write_json(capture_path, capture)

    # Existing projections contain hashes tied to the original capture and to
    # the original verifier path.  They are regenerated in the isolated copy.
    for name in ("summary.json", "allocation-metrics.json"):
        (bundle / name).unlink(missing_ok=True)


def run_verifier(bundle: Path, *, write: bool) -> subprocess.CompletedProcess[str]:
    argv = [sys.executable, str(bundle / "verify.py"), "--root", str(bundle)]
    if write:
        argv.append("--write")
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    environment["PYTHONPATH"] = str(bundle.parent)
    return subprocess.run(
        argv,
        cwd=bundle.parent,
        env=environment,
        capture_output=True,
        text=True,
        timeout=600,
    )


def compress_one_report(bundle: Path, relative: str) -> bool:
    path = bundle / relative
    if not path.is_file():
        return any(Path(str(path) + suffix).is_file() for suffix in (".zst", ".zstd"))
    process = subprocess.run(
        ["zstd", "-q", "-c", str(path)],
        capture_output=True,
        check=False,
        timeout=120,
    )
    if process.returncode != 0:
        detail = process.stderr.decode("utf-8", "replace").strip()
        raise RuntimeError(f"cannot create compressed report: {detail}")
    path.with_name(path.name + ".zst").write_bytes(process.stdout)
    path.unlink()
    return True


def prepare_export(source: Path, temporary: Path, repo_root: Path) -> tuple[Path, list[str]]:
    bundle = copy_source_bundle(source, temporary, repo_root)
    portableize(bundle, source)
    compressed: list[str] = []
    for phase, selector in (
        ("normal", "pptx_cross_copy_media_rich_lifecycle"),
        ("allocator", "pptx_cross_copy_media_rich_lifecycle"),
    ):
        relative = formal_report(phase, selector)
        if compress_one_report(bundle, relative):
            compressed.append(relative)
    return bundle, compressed


def require_formal_report(bundle: Path, phase: str, selector: str) -> str:
    relative = formal_report(phase, selector)
    artifact_storage(bundle, relative)
    return relative


def run_probe(
    base: Path,
    temporary: Path,
    label: str,
    mutation: Callable[[Path], None],
    diagnostic_marker: str,
) -> dict[str, Any]:
    probe = temporary / f"probe-{label}"
    shutil.copytree(base, probe)
    try:
        mutation(probe)
        result = run_verifier(probe, write=True)
        diagnostic = result.stderr.strip().splitlines()
        if result.returncode == 0:
            raise RuntimeError(f"{label}: verifier unexpectedly accepted corruption")
        if not any(line.startswith("0418 verification failed:") for line in diagnostic):
            raise RuntimeError(
                f"{label}: verifier did not return a bounded failure: {result.stderr!r}"
            )
        if diagnostic_marker not in result.stderr:
            raise RuntimeError(
                f"{label}: failure did not identify {diagnostic_marker!r}: "
                f"{result.stderr!r}"
            )
        return {
            "mutation": label,
            "exit_code": result.returncode,
            "expected": "rejected",
            "diagnostic_marker": diagnostic_marker,
            "diagnostic": diagnostic[-1] if diagnostic else "",
        }
    finally:
        shutil.rmtree(probe, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=SCRIPT_ROOT)
    parser.add_argument(
        "--output",
        type=Path,
        default=SCRIPT_ROOT / "checks" / "negative-probes.json",
    )
    args = parser.parse_args()
    source = args.root.resolve()
    repo_root = find_repo_root()
    if not (source / "capture.json").is_file():
        raise RuntimeError(f"capture is missing from {source}")
    if shutil.which("zstd") is None:
        raise RuntimeError("zstd is required for the compressed-report fallback probe")

    probes: list[dict[str, Any]] = []
    compressed_reports: list[str] = []
    with tempfile.TemporaryDirectory(prefix="litchi-goal-0418-negative-") as td:
        temporary = Path(td)
        # First exercise the source export exactly as supplied, including its
        # retained projections.  The second copy below is the portable,
        # path-rewritten export whose projections are regenerated in isolation.
        unmodified = copy_source_bundle(source, temporary / "unmodified", repo_root)
        unmodified_result = run_verifier(unmodified, write=False)
        if unmodified_result.returncode != 0:
            raise RuntimeError(
                "unmodified isolated export replay failed: "
                f"{unmodified_result.stdout} {unmodified_result.stderr}"
            )

        base, compressed_reports = prepare_export(source, temporary, repo_root)
        written = run_verifier(base, write=True)
        if written.returncode != 0:
            raise RuntimeError(
                f"isolated export write failed: {written.stdout} {written.stderr}"
            )
        replay = run_verifier(base, write=False)
        if replay.returncode != 0:
            raise RuntimeError(
                f"isolated export replay failed: {replay.stdout} {replay.stderr}"
            )

        lifecycle_normal = formal_report(
            "normal", "pptx_cross_copy_media_rich_lifecycle"
        )
        lifecycle_allocator = formal_report(
            "allocator", "pptx_cross_copy_media_rich_lifecycle"
        )
        phase_normal = formal_report("normal", "pptx_cross_copy_media_rich")
        require_formal_report(base, "normal", "pptx_cross_copy_media_rich_lifecycle")
        require_formal_report(base, "allocator", "pptx_cross_copy_media_rich_lifecycle")
        require_formal_report(base, "normal", "pptx_cross_copy_media_rich")

        def report_mutation(
            relative: str, change: Callable[[dict[str, Any]], None]
        ) -> Callable[[Path], None]:
            return lambda root: mutate_json_artifact(root, relative, change)

        probes.append(
            run_probe(
                base,
                temporary,
                "normal-raw-vector",
                report_mutation(
                    lifecycle_normal,
                    lambda value: value["results"][0]["source"][
                        "pptx_cross_copy"
                    ]["plan_ns"].pop(),
                ),
                "plan_ns",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "allocator-raw-vector",
                report_mutation(
                    lifecycle_allocator,
                    lambda value: value["results"][0]["operation_metrics"][
                        "allocation"
                    ]["allocated_bytes"]["values"].pop(),
                ),
                "allocated_bytes",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "sample-order",
                report_mutation(
                    lifecycle_normal,
                    lambda value: value["results"][0]["elapsed_ns"][
                        "sample_order"
                    ].__setitem__(1, value["results"][0]["elapsed_ns"]["sample_order"][0]),
                ),
                "sample_order",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "phase-order",
                report_mutation(
                    phase_normal,
                    lambda value: value["results"][0]["source"][
                        "pptx_cross_copy"
                    ]["plan_ns"].__setitem__(
                        0,
                        value["results"][0]["source"]["pptx_cross_copy"][
                            "plan_ns"
                        ][0]
                        + 1,
                    ),
                ),
                "plan/commit/publication",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "source",
                report_mutation(
                    lifecycle_normal,
                    lambda value: value["results"][0]["source"][
                        "pptx_cross_copy"
                    ].__setitem__("source_archive_sha256", "0" * 64),
                ),
                "source identity",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "binary",
                report_mutation(
                    lifecycle_normal,
                    lambda value: value["binary_identity"].__setitem__(
                        "binary_sha256", "0" * 64
                    ),
                ),
                "binary_identity",
            )
        )

        def wrong_argv(value: dict[str, Any]) -> None:
            for run in value["runs"]:
                index = run["argv"].index("--warmup") + 1
                run["argv"][index] = "999"
                return
            raise RuntimeError("capture has no argv")

        probes.append(
            run_probe(
                base,
                temporary,
                "argv",
                lambda root: mutate_capture(root, wrong_argv),
                "capture.runs[0].argv",
            )
        )
        probes.append(
            run_probe(
                base,
                temporary,
                "oracle",
                report_mutation(
                    lifecycle_normal,
                    lambda value: value["results"][0].__setitem__(
                        "output_sha256", "0" * 64
                    ),
                ),
                "output_sha256",
            )
        )

        def false_gate(value: dict[str, Any]) -> None:
            gates = value["results"][0]["source"]["pptx_cross_copy"]["gates"]
            gates[next(iter(gates))] = False

        probes.append(
            run_probe(
                base,
                temporary,
                "gate",
                report_mutation(lifecycle_normal, false_gate),
                "gates",
            )
        )

        catalog = "normal/a1-pptx_cross_copy_media_rich_lifecycle.catalog.json"
        probes.append(
            run_probe(
                base,
                temporary,
                "catalog",
                report_mutation(
                    catalog,
                    lambda value: value.__setitem__("catalog_sha256", "0" * 64),
                ),
                "corpus_catalog",
            )
        )

        def zero_rss(text: str) -> str:
            marker = "Maximum resident set size (kbytes):"
            lines = text.splitlines(keepends=True)
            changed = 0
            for index, line in enumerate(lines):
                if marker in line:
                    prefix, suffix = line.split(":", 1)
                    ending = "\n" if line.endswith("\n") else ""
                    lines[index] = prefix + ": 0" + ending
                    changed += 1
            if changed != 1:
                raise RuntimeError(f"expected one RSS field, found {changed}")
            return "".join(lines)

        probes.append(
            run_probe(
                base,
                temporary,
                "time-rss",
                lambda root: mutate_text_artifact(
                    root,
                    "preflight/a1-normal-pptx_cross_copy_media_rich_lifecycle.time.txt",
                    zero_rss,
                ),
                "maximum RSS",
            )
        )

        def bad_artifact_hash(value: dict[str, Any]) -> None:
            value["runs"][0]["artifact_sha256"]["report"] = "0" * 64

        probes.append(
            run_probe(
                base,
                temporary,
                "artifacthash",
                lambda root: mutate_capture(root, bad_artifact_hash),
                "artifact_sha256",
            )
        )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    output = {
        "status": "passed",
        "unmodified_isolated_export_replay": True,
        "path_rewritten_isolated_export_replays": 2,
        "compressed_report_fallback": compressed_reports,
        "temporary_export_removed": True,
        "probes": probes,
    }
    args.output.write_text(
        json.dumps(output, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "status": "passed",
                "unmodified_isolated_export_replay": True,
                "path_rewritten_isolated_export_replays": 2,
                "corruptions_rejected": len(probes),
                "compressed_reports": compressed_reports,
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, subprocess.SubprocessError) as error:
        print(f"0418 negative probes failed: {error}", file=sys.stderr)
        raise SystemExit(1)
