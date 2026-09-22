#!/usr/bin/env python3
"""Run offline corruption controls through the 0734 analyzer and audit.

Every control starts with a complete packet copy, rebases packet-local paths,
and mutates one evidence field.  Immediate dependent digests are refreshed
when the control is intended to reach a deeper validator.  The real packet,
workspace, fixture files, binaries, and captures are never changed.
"""

from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import inspect
import io
import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
sys.dont_write_bytecode = True


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def copy_packet(target: Path) -> None:
    shutil.copytree(
        P,
        target,
        ignore=shutil.ignore_patterns(
            "__pycache__", "preflight-attempt-*", "negative-attempt-*"
        ),
    )


def load_preflight() -> Any:
    spec = importlib.util.spec_from_file_location(
        "change0734_negative_preflight", P / "preflight.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load preflight helpers")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def rebase_receipts(packet: Path, helper: Any) -> None:
    """Rebind copied preflight and capture manifest digests."""
    helper.rebase_packet(packet)
    receipt_path = packet / "preflight.json"
    receipt = read(receipt_path)
    receipt["freeze_sha256"] = sha(packet / "freeze.json")
    scripts = receipt.get("scripts")
    if isinstance(scripts, dict):
        for name in list(scripts):
            copied = packet / name
            if copied.is_file():
                scripts[name] = sha(copied)
    for field, name in (
        ("script_sha256", "preflight.py"),
        ("analyzer_sha256", "analyze.py"),
        ("audit_sha256", "audit.py"),
        ("auditor_sha256", "audit.py"),
    ):
        if field in receipt:
            receipt[field] = sha(packet / name)
    write(receipt_path, receipt)
    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    manifest["freeze_sha256"] = sha(packet / "freeze.json")
    manifest["preflight_sha256"] = sha(receipt_path)
    write(manifest_path, manifest)


def freeze_map(packet: Path) -> tuple[dict[str, str], bool]:
    frozen = read(packet / "freeze.json")
    if isinstance(frozen.get("bindings"), dict):
        return frozen["bindings"], True
    return frozen, False


def refresh_freeze(packet: Path, helper: Any) -> None:
    """Refresh copied packet bindings after a deliberate deep mutation."""
    bindings, wrapped = freeze_map(packet)
    updated: dict[str, str] = {}
    for raw, expected in bindings.items():
        path = Path(raw)
        updated[raw] = sha(path) if path.is_file() else expected
    frozen = read(packet / "freeze.json")
    if wrapped:
        frozen["bindings"] = updated
    else:
        frozen = updated
    write(packet / "freeze.json", frozen)


def refresh_dependent_custody(packet: Path, helper: Any) -> None:
    """Refresh hashes needed to expose a mutation below packet custody."""
    for build_path in sorted(packet.glob("*-build.json")):
        build = read(build_path)
        quality_name = build.get("quality")
        if isinstance(quality_name, str):
            quality_path = packet / quality_name
            if quality_path.is_file():
                build["quality_sha256"] = sha(quality_path)
        write(build_path, build)
    qualification_path = packet / "qualification.json"
    qualification = read(qualification_path)
    files = qualification.get("files")
    if isinstance(files, dict):
        for raw in list(files):
            path = packet / raw
            if path.is_file():
                files[raw] = sha(path)
    builds = qualification.get("builds")
    if isinstance(builds, dict):
        for variant in list(builds):
            path = packet / f"{variant}-build.json"
            if path.is_file():
                builds[variant] = sha(path)
    write(qualification_path, qualification)
    refresh_freeze(packet, helper)
    rebase_receipts(packet, helper)


def configure(module: Any, packet: Path) -> None:
    for name in ("P", "PACKET", "ROOT"):
        if hasattr(module, name):
            setattr(module, name, packet if name != "ROOT" else ROOT)
    if hasattr(module, "CAPTURES"):
        module.CAPTURES = packet / "captures"


def load_analyzer(packet: Path) -> Any:
    spec = importlib.util.spec_from_file_location(
        "change0734_negative_analyzer", packet / "analyze.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load copied analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    configure(module, packet)
    return module


def load_audit(packet: Path) -> Any:
    spec = importlib.util.spec_from_file_location(
        "change0734_negative_audit", packet / "audit.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load copied audit")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    configure(module, packet)
    return module


def run_validator(module: Any, packet: Path) -> tuple[bool, str | None]:
    configure(module, packet)
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            module.main()
    except Exception as error:
        failure = getattr(module, "Failure", None)
        if failure is not None and isinstance(error, failure):
            return False, f"{type(error).__name__}: {error}"
        raise
    return True, None


def run_audit(module: Any, packet: Path) -> tuple[bool, str | None]:
    configure(module, packet)
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            module.main()
    except Exception as error:
        failure = getattr(module, "AuditError", None)
        if failure is not None and isinstance(error, failure):
            return False, f"{type(error).__name__}: {error}"
        raise
    return True, None


def rows(packet: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    manifest = read(packet / "captures" / "manifest.json")
    runs = manifest["runs"]
    row = next(
        row for row in runs
        if row["lane"] == "native"
        and row["case"] == "primary"
        and row["variant"] == "baseline"
    )
    alloc = next(
        row for row in runs
        if row["lane"] == "allocation"
        and row["case"] == "primary"
        and row["variant"] == "baseline"
    )
    return row, alloc


def report_path(packet: Path, row: dict[str, Any]) -> Path:
    return packet / "captures" / row["output"]


def edit_report(
    packet: Path,
    mutate: Callable[[dict[str, Any]], None],
    *,
    allocation: bool = False,
    update_digest: bool = True,
) -> None:
    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    lane = "allocation" if allocation else "native"
    row = next(
        row for row in manifest["runs"]
        if row["lane"] == lane
        and row["case"] == "primary"
        and row["variant"] == "baseline"
    )
    output = report_path(packet, row)
    report = read(output)
    mutate(report)
    write(output, report)
    if update_digest:
        row["sha256"] = sha(output)
    write(manifest_path, manifest)


def edit_sample(
    packet: Path,
    mutate: Callable[[dict[str, Any]], None],
    *,
    allocation: bool = False,
) -> None:
    edit_report(
        packet,
        lambda report: mutate(report["samples"][0]),
        allocation=allocation,
    )


def edit_manifest(packet: Path, mutate: Callable[[dict[str, Any]], None]) -> None:
    path = packet / "captures" / "manifest.json"
    manifest = read(path)
    mutate(manifest)
    write(path, manifest)


def update_manifest_freeze_digest(packet: Path) -> None:
    path = packet / "captures" / "manifest.json"
    manifest = read(path)
    manifest["freeze_sha256"] = sha(packet / "freeze.json")
    write(path, manifest)


def update_manifest_preflight_digest(packet: Path) -> None:
    path = packet / "captures" / "manifest.json"
    manifest = read(path)
    manifest["preflight_sha256"] = sha(packet / "preflight.json")
    write(path, manifest)


def semantic_output_hash(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample.__setitem__("output_sha256", "0" * 64))


def semantic_oracle(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample["oracle"].__setitem__("semantic_reopen_ok", False))


def semantic_control(packet: Path) -> None:
    edit_report(packet, lambda report: report["oracle_controls"].pop())


def source_hash(packet: Path) -> None:
    edit_report(packet, lambda report: report.__setitem__("source_sha256", "0" * 64))


def allocation_shape(packet: Path) -> None:
    def mutate(sample: dict[str, Any]) -> None:
        sample["allocations"]["whole"]["allocated_bytes"] = -1

    edit_sample(packet, mutate, allocation=True)


def allocation_missing(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample.pop("allocations"), allocation=True)


def reordered_schedule(packet: Path) -> None:
    def mutate(manifest: dict[str, Any]) -> None:
        manifest["runs"][0], manifest["runs"][1] = manifest["runs"][1], manifest["runs"][0]

    edit_manifest(packet, mutate)


def missing_schedule_row(packet: Path) -> None:
    edit_manifest(packet, lambda manifest: manifest["runs"].pop())


def wrong_command(packet: Path) -> None:
    edit_manifest(packet, lambda manifest: manifest["runs"][0]["command"].__setitem__(-1, "999"))


def raw_report_digest(packet: Path) -> None:
    row, _ = rows(packet)
    path = report_path(packet, row)
    path.write_bytes(path.read_bytes() + b"\nnegative-control\n")


def raw_stderr_digest(packet: Path) -> None:
    row, _ = rows(packet)
    path = packet / "captures" / (row["output"] + ".stderr")
    path.write_bytes(path.read_bytes() + b"negative-control\n")


def freeze_digest(packet: Path) -> None:
    frozen = read(packet / "freeze.json")
    bindings, wrapped = freeze_map(packet)
    target = next(key for key in bindings if key.endswith("/plan.json"))
    bindings[target] = "0" * 64
    if wrapped:
        frozen["bindings"] = bindings
    else:
        frozen = bindings
    write(packet / "freeze.json", frozen)
    update_manifest_freeze_digest(packet)


def preflight_digest(packet: Path) -> None:
    receipt = read(packet / "preflight.json")
    receipt["scripts"]["analyze.py"] = "0" * 64
    write(packet / "preflight.json", receipt)
    update_manifest_preflight_digest(packet)


def build_source_digest(packet: Path, helper: Any) -> None:
    path = packet / "candidate-build.json"
    build = read(path)
    source = next(iter(build["source"]))
    build["source"][source] = "0" * 64
    write(path, build)
    refresh_dependent_custody(packet, helper)


def quality_command(packet: Path, helper: Any) -> None:
    build = read(packet / "candidate-build.json")
    quality_path = packet / build["quality"]
    quality = read(quality_path)
    quality["runs"][0]["command"][0] = "cargo-corrupted"
    write(quality_path, quality)
    refresh_dependent_custody(packet, helper)


def quality_output_digest(packet: Path) -> None:
    build = read(packet / "candidate-build.json")
    quality = read(packet / build["quality"])
    output = packet / quality["runs"][0]["output"]
    output.write_bytes(output.read_bytes() + b"\nnegative-control\n")


def qualification_digest(packet: Path, helper: Any) -> None:
    path = packet / "qualification.json"
    qualification = read(path)
    key = "qualification-raw/baseline-primary-native.json"
    assert key in qualification["files"]
    qualification["files"][key] = "0" * 64
    write(path, qualification)
    # Rebind the edited receipt and its freeze entry without repairing the
    # deliberately false file-map digest.  This reaches the qualification
    # validator after packet-level custody has passed.
    refresh_freeze(packet, helper)
    receipt = read(packet / "preflight.json")
    receipt["freeze_sha256"] = sha(packet / "freeze.json")
    write(packet / "preflight.json", receipt)
    manifest = read(packet / "captures" / "manifest.json")
    manifest["freeze_sha256"] = sha(packet / "freeze.json")
    manifest["preflight_sha256"] = sha(packet / "preflight.json")
    write(packet / "captures" / "manifest.json", manifest)


def candidate_binary_digest(packet: Path, helper: Any) -> None:
    build = read(packet / "candidate-build.json")
    binary = build["binaries"]["native"]
    # Keep the binary bytes untouched; alter the retained receipt and refresh
    # only the copied custody chain so the binary identity check is exercised.
    binary["sha256"] = "0" * 64
    write(packet / "candidate-build.json", build)
    refresh_dependent_custody(packet, helper)


def report_route(packet: Path) -> None:
    edit_report(packet, lambda report: report.__setitem__("case", "ppt-secondary"))


Control = tuple[
    str,
    str,
    str,
    Callable[[Path, Any], None] | Callable[[Path], None],
]


def invoke(mutate: Callable[..., None], packet: Path, helper: Any) -> None:
    if len(inspect.signature(mutate).parameters) >= 2:
        mutate(packet, helper)
    else:
        mutate(packet)


CONTROLS: tuple[Control, ...] = (
    ("semantic output hash", "semantic", "output identity changed", semantic_output_hash),
    ("semantic oracle witness", "semantic", "semantic oracle changed", semantic_oracle),
    ("missing semantic control", "semantic-controls", "oracle_controls changed", semantic_control),
    ("source hash", "semantic-custody", "static field source_sha256 changed", source_hash),
    ("report case identity", "semantic-identity", "static field case changed", report_route),
    ("allocation negative bytes", "allocation", "allocated_bytes", allocation_shape),
    ("allocation missing witness", "allocation", "allocation sample shape changed", allocation_missing),
    ("reordered schedule", "schedule", "capture schedule changed", reordered_schedule),
    ("missing schedule row", "schedule", "capture process count changed", missing_schedule_row),
    ("wrong capture command", "schedule", "capture command changed", wrong_command),
    ("raw report digest", "capture-custody", "capture output digest changed", raw_report_digest),
    ("raw stderr digest", "capture-custody", "capture stderr digest changed", raw_stderr_digest),
    ("frozen input digest", "freeze-custody", "frozen input changed", freeze_digest),
    ("preflight script digest", "preflight-custody", "preflight script binding changed", preflight_digest),
    ("build source digest", "build-custody", "candidate source changed", build_source_digest),
    ("candidate binary digest", "build-custody", "binary identity changed", candidate_binary_digest),
    ("quality command", "quality-custody", "quality command 0 changed", quality_command),
    ("quality output digest", "quality-custody", "quality output 0 digest changed", quality_output_digest),
    ("qualification output digest", "qualification-custody", "candidate qualification file changed", qualification_digest),
)


def normalize_error(error: str | None, packet: Path) -> str | None:
    if error is None:
        return None
    return (
        error.replace(str(packet), "<packet>")
        .replace(str(P), "<packet-source>")
        .replace(str(ROOT), "<workspace>")
    )


def main() -> None:
    assert (P / "preflight.json").is_file(), "run preflight.py before negative checks"
    helper = load_preflight()
    checks: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="litchi-0734-negative-") as temporary:
        temporary_root = Path(temporary) / "isolated" / "results"
        temporary_root.mkdir(parents=True)

        baseline_packet = temporary_root / "baseline"
        copy_packet(baseline_packet)
        helper.rebase_packet(baseline_packet)
        rebase_receipts(baseline_packet, helper)
        module = load_analyzer(baseline_packet)
        audit_module = load_audit(baseline_packet)
        accepted, error = run_validator(module, baseline_packet)
        if not accepted:
            raise AssertionError(f"complete packet rejected: {error}")
        audit_accepted, audit_error = run_audit(audit_module, baseline_packet)
        if not audit_accepted:
            raise AssertionError(f"complete packet audit rejected: {audit_error}")
        checks.append({"name": "complete packet accepted", "accepted": True, "expected": True})

        for name, category, expected_fragment, mutate in CONTROLS:
            packet = temporary_root / ("control-" + str(len(checks)))
            copy_packet(packet)
            helper.rebase_packet(packet)
            rebase_receipts(packet, helper)
            invoke(mutate, packet, helper)
            analyzer_accepted, analyzer_error = run_validator(module, packet)
            audit_accepted, audit_error = run_audit(audit_module, packet)
            if analyzer_accepted and audit_accepted:
                raise AssertionError(f"control was accepted: {name}")
            errors = (analyzer_error or "") + "\n" + (audit_error or "")
            if expected_fragment not in errors:
                raise AssertionError(
                    f"control {name} rejected for an unexpected reason: {errors!r}; "
                    f"expected {expected_fragment!r}"
                )
            checks.append({
                "name": name,
                "category": category,
                "expected_rejection": expected_fragment,
                "accepted": False,
                "expected": False,
                "analyzer_accepted": analyzer_accepted,
                "analyzer_rejection": normalize_error(analyzer_error, packet),
                "audit_accepted": audit_accepted,
                "audit_rejection": normalize_error(audit_error, packet),
                "rejection_coverage": [
                    name for name, accepted in
                    (("analyzer", analyzer_accepted), ("audit", audit_accepted))
                    if not accepted
                ],
            })

    receipt = {
        "status": "passed",
        "kind": "actual-analyzer-corruption-controls",
        "control_count": len(CONTROLS),
        "rejection_count": sum(not item["accepted"] for item in checks[1:]),
        "checks": checks,
        "analyzer_sha256": sha(P / "analyze.py"),
        "script_sha256": sha(Path(__file__).resolve()),
    }
    write(P / "negative-checks.json", receipt)
    print(f"PASS actual analyzer baseline plus {len(CONTROLS)} independent rejection controls")


if __name__ == "__main__":
    main()
