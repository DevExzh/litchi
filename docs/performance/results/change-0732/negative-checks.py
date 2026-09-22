#!/usr/bin/env python3
"""Run corruption controls through the real 0732 analyzer in temp packets.

The controls are deliberately offline.  Each one starts with a complete copy
of the packet, rebases packet-local absolute paths, and mutates one evidence
field.  The workspace, sealed packet, native binary, and capture files are
never changed.
"""

from __future__ import annotations

import contextlib
import hashlib
import importlib.util
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
    """Copy the packet, retaining captures and excluding Python caches."""
    shutil.copytree(
        P,
        target,
        ignore=shutil.ignore_patterns("__pycache__"),
    )


def replace_packet_prefix(value: Any, original: str, temporary: str) -> Any:
    if isinstance(value, str):
        if value == original:
            return temporary
        prefix = original + "/"
        if value.startswith(prefix):
            return temporary + value[len(original):]
        return value
    if isinstance(value, list):
        return [replace_packet_prefix(item, original, temporary) for item in value]
    if isinstance(value, dict):
        return {
            key: replace_packet_prefix(item, original, temporary)
            for key, item in value.items()
        }
    return value


def rebase_quality(packet: Path) -> None:
    """Rebase quality command paths and refresh dependent synthetic custody."""
    original = str(P)
    temporary = str(packet)
    build_path = packet / "build.json"
    build = read(build_path)
    quality_path = packet / build["quality"]
    quality = replace_packet_prefix(read(quality_path), original, temporary)
    write(quality_path, quality)
    build["quality_sha256"] = sha(quality_path)
    write(build_path, build)

    qualification_path = packet / "qualification.json"
    qualification = read(qualification_path)
    qualification_manifest_path = packet / "qualification" / "manifest.json"
    if qualification_manifest_path.is_file():
        qualification_manifest = replace_packet_prefix(
            read(qualification_manifest_path), original, temporary
        )
        write(qualification_manifest_path, qualification_manifest)
        files = qualification.get("files")
        if isinstance(files, dict) and "qualification/manifest.json" in files:
            files["qualification/manifest.json"] = sha(qualification_manifest_path)
    qualification["build_sha256"] = sha(build_path)
    write(qualification_path, qualification)


def rebase_freeze(packet: Path) -> None:
    """Rebase packet-owned freeze keys and hash their copied targets."""
    original = str(P)
    temporary = str(packet)
    frozen = read(packet / "freeze.json")
    if isinstance(frozen.get("bindings"), dict):
        bindings = frozen["bindings"]
        wrapped = True
    else:
        bindings = frozen
        wrapped = False

    rebased: dict[str, str] = {}
    for raw, expected in bindings.items():
        key = replace_packet_prefix(raw, original, temporary)
        target = Path(key)
        # The measured binary and workspace inputs remain at their real paths;
        # copied packet paths have their copied bytes hashed below.
        rebased[key] = sha(target) if target.is_file() else expected
    if wrapped:
        frozen["bindings"] = rebased
    else:
        frozen = rebased
    write(packet / "freeze.json", frozen)


def rebase_preflight(packet: Path) -> None:
    """Rebind the copied preflight receipt and capture's preflight digest."""
    receipt_path = packet / "preflight.json"
    if not receipt_path.is_file():
        raise RuntimeError("preflight.json is required for 0732 negative checks")
    receipt = read(receipt_path)
    receipt["freeze_sha256"] = sha(packet / "freeze.json")
    scripts = receipt.get("scripts")
    if isinstance(scripts, dict):
        for name in scripts:
            script_path = packet / name
            if script_path.is_file():
                scripts[name] = sha(script_path)
    # These fields are retained in the current receipt for convenient human
    # inspection; update them when present without requiring them in the
    # analyzer contract.
    if "script_sha256" in receipt:
        receipt["script_sha256"] = sha(packet / "preflight.py")
    if "analyzer_sha256" in receipt:
        receipt["analyzer_sha256"] = sha(packet / "analyze.py")
    if "audit_sha256" in receipt:
        receipt["audit_sha256"] = sha(packet / "audit.py")
    if "auditor_sha256" in receipt:
        receipt["auditor_sha256"] = sha(packet / "audit.py")
    write(receipt_path, receipt)

    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    manifest["freeze_sha256"] = sha(packet / "freeze.json")
    manifest["preflight_sha256"] = sha(receipt_path)
    write(manifest_path, manifest)


def prepare(packet: Path) -> None:
    # Rebase freeze first so quality/build changes can refresh their copied
    # entries.  The receipt is generated after that freeze has its final hash.
    rebase_freeze(packet)
    rebase_quality(packet)
    rebase_freeze(packet)
    rebase_preflight(packet)


def load_analyzer(packet: Path):
    spec = importlib.util.spec_from_file_location(
        "change0732_negative_analyzer", packet / "analyze.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load copied analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.P = packet
    module.ROOT = ROOT
    module.CAPTURES = packet / "captures"
    return module


def capture_row(packet: Path, route: str = "profiled-clock") -> tuple[dict[str, Any], Path, dict[str, Any]]:
    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    row = next(row for row in manifest["runs"] if row["route"] == route)
    output = packet / "captures" / row["output"]
    return row, output, manifest


def edit_report(
    packet: Path,
    mutate: Callable[[dict[str, Any]], None],
    route: str = "profiled-clock",
    update_digest: bool = True,
) -> None:
    row, output, manifest = capture_row(packet, route)
    report = read(output)
    mutate(report)
    write(output, report)
    if update_digest:
        row["sha256"] = sha(output)
    write(packet / "captures" / "manifest.json", manifest)


def edit_sample(packet: Path, mutate: Callable[[dict[str, Any]], None], route: str = "profiled-clock") -> None:
    edit_report(packet, lambda report: mutate(report["samples"][0]), route)


def edit_diagnostic(packet: Path, mutate: Callable[[dict[str, Any]], None]) -> None:
    def apply(report: dict[str, Any]) -> None:
        mutate(report["samples"][0]["diagnostics"]["commit"])

    edit_report(packet, apply)


def edit_manifest(packet: Path, mutate: Callable[[dict[str, Any]], None]) -> None:
    path = packet / "captures" / "manifest.json"
    manifest = read(path)
    mutate(manifest)
    write(path, manifest)


def edit_freeze_digest(packet: Path) -> None:
    path = packet / "freeze.json"
    frozen = read(path)
    bindings = frozen.get("bindings", frozen)
    target = next(key for key in bindings if key.endswith("/plan.json"))
    bindings[target] = "0" * 64
    if "bindings" in frozen:
        frozen["bindings"] = bindings
    else:
        frozen = bindings
    write(path, frozen)
    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    manifest["freeze_sha256"] = sha(path)
    write(manifest_path, manifest)


def edit_build_digest(packet: Path) -> None:
    path = packet / "build.json"
    build = read(path)
    source = next(iter(build["source"]))
    build["source"][source] = "0" * 64
    write(path, build)


def edit_quality_digest(packet: Path) -> None:
    build = read(packet / "build.json")
    quality_path = packet / build["quality"]
    quality = read(quality_path)
    output = packet / quality["runs"][0]["output"]
    # Leave the manifest's recorded digest unchanged; verify_quality must
    # reject the changed retained command output as a custody failure.
    output.write_bytes(output.read_bytes() + b"\n")


def edit_preflight_digest(packet: Path) -> None:
    receipt_path = packet / "preflight.json"
    receipt = read(receipt_path)
    receipt["scripts"]["analyze.py"] = "0" * 64
    write(receipt_path, receipt)
    manifest_path = packet / "captures" / "manifest.json"
    manifest = read(manifest_path)
    manifest["preflight_sha256"] = sha(receipt_path)
    write(manifest_path, manifest)


def edit_raw_report(packet: Path) -> None:
    _row, output, _manifest = capture_row(packet)
    output.write_bytes(output.read_bytes() + b"\n")


def edit_raw_stderr(packet: Path) -> None:
    row, _output, _manifest = capture_row(packet)
    stderr = packet / "captures" / (row["output"] + ".stderr")
    stderr.write_bytes(stderr.read_bytes() + b"negative-control\n")


def event_missing(packet: Path) -> None:
    edit_diagnostic(packet, lambda report: report["events"].pop(0))


def event_reordered(packet: Path) -> None:
    def mutate(report: dict[str, Any]) -> None:
        report["events"][0], report["events"][1] = report["events"][1], report["events"][0]

    edit_diagnostic(packet, mutate)


def event_failed(packet: Path) -> None:
    edit_diagnostic(packet, lambda report: report["events"][1].__setitem__("outcome", "error"))


def event_nonmonotonic(packet: Path) -> None:
    def mutate(report: dict[str, Any]) -> None:
        report["events"][0]["t_ns"] = report["events"][1]["t_ns"] + 1

    edit_diagnostic(packet, mutate)


def event_outside_window(packet: Path) -> None:
    def mutate(sample: dict[str, Any]) -> None:
        report = sample["diagnostics"]["commit"]
        outside = sample["split"]["commit_end_ns"] + 1
        # Move the final zero-duration event pair just outside the commit
        # window.  This preserves monotonicity, event/span correspondence, and
        # the total-duration bound so the window check is the first failure.
        report["events"][-2]["t_ns"] = outside
        report["events"][-1]["t_ns"] = outside
        span = report["spans"][-1]
        span["start_ns"] = outside
        span["finish_ns"] = outside
        span["duration_ns"] = 0

    edit_sample(packet, mutate)


def span_arithmetic(packet: Path) -> None:
    edit_diagnostic(packet, lambda report: report["spans"][0].__setitem__(
        "duration_ns", report["spans"][0]["duration_ns"] + 1
    ))


def split_arithmetic(packet: Path) -> None:
    def mutate(sample: dict[str, Any]) -> None:
        sample["split"]["split_sum_ns"] += 1

    edit_sample(packet, mutate, "profiled-empty")


def zero_whole(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample.__setitem__("whole_ns", 0), "ordinary-opaque")


def false_oracle(packet: Path) -> None:
    def mutate(sample: dict[str, Any]) -> None:
        sample["oracle"]["oracle_ok"] = False

    edit_sample(packet, mutate, "ordinary-opaque")


def missing_control(packet: Path) -> None:
    edit_report(packet, lambda report: report["oracle_controls"].pop(), "ordinary-opaque")


def source_hash(packet: Path) -> None:
    edit_report(packet, lambda report: report.__setitem__("source_sha256", "0" * 64), "ordinary-opaque")


def report_route(packet: Path) -> None:
    edit_report(packet, lambda report: report.__setitem__("route", "ordinary-split"), "ordinary-opaque")


def sample_route(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample.__setitem__("route", "ordinary-split"), "ordinary-opaque")


def output_hash(packet: Path) -> None:
    edit_sample(packet, lambda sample: sample.__setitem__("output_sha256", "0" * 64), "ordinary-opaque")


def wrong_command(packet: Path) -> None:
    edit_manifest(packet, lambda manifest: manifest["runs"][0]["command"].__setitem__(-1, "999"))


def reordered_schedule(packet: Path) -> None:
    def mutate(manifest: dict[str, Any]) -> None:
        manifest["runs"][0], manifest["runs"][1] = manifest["runs"][1], manifest["runs"][0]

    edit_manifest(packet, mutate)


def missing_row(packet: Path) -> None:
    edit_manifest(packet, lambda manifest: manifest["runs"].pop())


Control = tuple[str, str, str, Callable[[Path], None]]


CONTROLS: tuple[Control, ...] = (
    ("missing diagnostic event", "diagnostic-sequence", "event count changed", event_missing),
    ("reordered diagnostic event", "diagnostic-sequence", "sequence/outcome changed", event_reordered),
    ("failed diagnostic outcome", "diagnostic-outcome", "sequence/outcome changed", event_failed),
    ("nonmonotonic diagnostic timestamp", "diagnostic-timestamp", "timestamps are not monotonic", event_nonmonotonic),
    ("event outside commit window", "diagnostic-window", "escapes commit window", event_outside_window),
    ("span duration arithmetic", "diagnostic-arithmetic", "span 0 duration changed", span_arithmetic),
    ("split sum arithmetic", "split-arithmetic", "split_sum_ns differs from phase sum", split_arithmetic),
    ("zero whole duration", "timing-arithmetic", "whole_ns is invalid", zero_whole),
    ("false semantic oracle", "semantic-oracle", "semantic oracle changed", false_oracle),
    ("missing corruption control", "semantic-controls", "oracle_controls differs", missing_control),
    ("source hash", "report-custody", "source_sha256 differs", source_hash),
    ("wrong report route", "route-identity", "report route differs from manifest", report_route),
    ("wrong sample route", "route-identity", "sample route changed", sample_route),
    ("output hash", "report-custody", "output digest changed", output_hash),
    ("wrong command", "capture-command", "command or exit status changed", wrong_command),
    ("reordered schedule", "capture-schedule", "changed schedule order", reordered_schedule),
    ("missing capture row", "capture-completeness", "capture process count changed", missing_row),
    ("raw report digest", "capture-custody", "capture output", edit_raw_report),
    ("raw stderr digest", "capture-custody", "capture stderr", edit_raw_stderr),
    ("frozen input digest", "freeze-custody", "frozen input changed", edit_freeze_digest),
    ("preflight script digest", "preflight-custody", "preflight script changed", edit_preflight_digest),
    ("build source digest", "build-custody", "build source", edit_build_digest),
    ("quality output digest", "quality-custody", "quality output", edit_quality_digest),
)


def run_one(module: Any, packet: Path) -> tuple[bool, str | None]:
    module.P = packet
    module.ROOT = ROOT
    module.CAPTURES = packet / "captures"
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            module.main()
    except module.Failure as error:
        return False, f"{type(error).__name__}: {error}"
    return True, None


def normalize_error(error: str | None, packet: Path) -> str | None:
    if error is None:
        return None
    return (
        error.replace(str(packet), "<packet>")
        .replace(str(P), "<packet-source>")
        .replace(str(ROOT), "<workspace>")
    )


def main() -> None:
    checks: list[dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="litchi-0732-negative-") as temporary:
        temporary_root = Path(temporary) / "isolated" / "results"
        temporary_root.mkdir(parents=True)

        baseline_packet = temporary_root / "baseline"
        copy_packet(baseline_packet)
        prepare(baseline_packet)
        module = load_analyzer(baseline_packet)
        accepted, error = run_one(module, baseline_packet)
        if not accepted:
            raise AssertionError(f"baseline packet rejected: {error}")
        checks.append({"name": "complete packet accepted", "accepted": True, "expected": True})

        for name, category, expected_fragment, mutate in CONTROLS:
            packet = temporary_root / ("control-" + str(len(checks)))
            copy_packet(packet)
            prepare(packet)
            mutate(packet)
            accepted, error = run_one(module, packet)
            if accepted:
                raise AssertionError(f"control was accepted: {name}")
            if expected_fragment not in (error or ""):
                raise AssertionError(
                    f"control {name} rejected for unexpected reason: {error!r}; "
                    f"expected {expected_fragment!r}"
                )
            checks.append({
                "name": name,
                "category": category,
                "expected_rejection": expected_fragment,
                "accepted": False,
                "expected": False,
                "rejection": normalize_error(error, packet),
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
