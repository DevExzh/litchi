#!/usr/bin/env python3
"""Exercise the 0732 analyzer and audit on synthetic full-matrix reports.

The preflight uses the four one-sample qualification reports as schemas and
copies each sample 50 times.  It never starts Cargo, the native probe, or a
profiler.  The temporary packet is discarded before this script returns, so
its output is an integration check rather than timing evidence.
"""

from __future__ import annotations

import contextlib
import copy
import hashlib
import importlib.util
import io
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
ROUTES = ("ordinary-opaque", "ordinary-split", "profiled-empty", "profiled-clock")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def copy_packet(target: Path) -> None:
    """Copy the packet while leaving native captures and Python caches out."""
    target.mkdir(exist_ok=True)
    for child in P.iterdir():
        if child.name in {"captures", "__pycache__"}:
            continue
        destination = target / child.name
        if child.is_dir():
            shutil.copytree(
                child,
                destination,
                ignore=shutil.ignore_patterns("__pycache__", "captures"),
            )
        elif child.is_file():
            shutil.copy2(child, destination)

    (target / "captures").mkdir()


def rebase_packet_freeze(packet: Path) -> None:
    """Rebase only packet-owned freeze keys; preserve workspace and binary paths."""
    original_prefix = str(P) + "/"
    frozen = read(packet / "freeze.json")
    if "bindings" in frozen:
        bindings = frozen["bindings"]
        wrapper = True
    else:
        bindings = frozen
        wrapper = False
    rebased = {}
    for raw, value in bindings.items():
        replacement = str(packet) + "/" + raw[len(original_prefix):] \
            if raw.startswith(original_prefix) else raw
        rebased[replacement] = value
    if wrapper:
        frozen["bindings"] = rebased
    else:
        frozen = rebased
    write(packet / "freeze.json", frozen)

    for raw in rebased:
        assert not raw.startswith(original_prefix), raw
    assert any(raw == str(ROOT / read(P / "case.json")["path"]) for raw in rebased)


def rebase_quality_manifest(packet: Path) -> None:
    """Adapt quality's packet-local manifest path for the isolated packet.

    The analyzer checks the exact command list.  Its expected path is based on
    the temporary packet root, while the retained quality manifest records the
    original packet root.  Updating this synthetic copy requires refreshing
    only the dependent temporary custody hashes; the real packet is untouched.
    """
    quality_name = read(packet / "build.json")["quality"]
    quality_path = packet / quality_name
    manifest = read(quality_path)
    original_prefix = str(P) + "/"
    temporary_prefix = str(packet) + "/"

    def replace(value: Any) -> Any:
        if isinstance(value, str) and value.startswith(original_prefix):
            return temporary_prefix + value[len(original_prefix):]
        if isinstance(value, list):
            return [replace(item) for item in value]
        if isinstance(value, dict):
            return {key: replace(item) for key, item in value.items()}
        return value

    rebased = replace(manifest)
    write(quality_path, rebased)

    # The copied build and qualification receipts bind the retained quality
    # manifest.  Refresh these synthetic-copy bindings and the corresponding
    # packet-owned freeze values so custody checks still run in full.
    build_path = packet / "build.json"
    build = read(build_path)
    build["quality_sha256"] = sha(quality_path)
    write(build_path, build)

    qualification_path = packet / "qualification.json"
    qualification = read(qualification_path)
    qualification["build_sha256"] = sha(build_path)
    write(qualification_path, qualification)

    freeze_path = packet / "freeze.json"
    frozen = read(freeze_path)
    if "bindings" in frozen:
        bindings = frozen["bindings"]
        wrapper = True
    else:
        bindings = frozen
        wrapper = False
    for path in (quality_path, build_path, qualification_path):
        bindings[str(path)] = sha(path)
    if wrapper:
        frozen["bindings"] = bindings
    else:
        frozen = bindings
    write(freeze_path, frozen)


def qualification_templates(packet: Path) -> dict[str, dict[str, Any]]:
    qualification = read(packet / "qualification/manifest.json")
    assert qualification.get("status") == "passed"
    rows = qualification.get("runs")
    assert isinstance(rows, list) and len(rows) == len(ROUTES)
    templates: dict[str, dict[str, Any]] = {}
    for row in rows:
        route = row.get("route")
        assert route in ROUTES and route not in templates
        report = read(packet / "qualification" / row["output"])
        assert isinstance(report.get("samples"), list) and len(report["samples"]) == 1
        templates[route] = report
    assert set(templates) == set(ROUTES)
    return templates


def expected_command(plan: dict[str, Any], build: dict[str, Any], case: dict[str, Any],
                     route: str) -> list[str]:
    return [
        "taskset", "-c", str(plan["cpu"]), build["binary"]["path"],
        "--route", route, "--input", case["path"],
        "--samples", str(plan["samples"]), "--warmups", str(plan["warmups"]),
    ]


def synthetic_captures(packet: Path) -> None:
    plan = read(packet / "plan.json")
    build = read(packet / "build.json")
    case = read(packet / "case.json")
    templates = qualification_templates(packet)
    captures = packet / "captures"
    runs = []

    for spec in plan["schedule"]:
        route = spec["route"]
        report = copy.deepcopy(templates[route])
        report["samples_requested"] = plan["samples"]
        report["warmups"] = plan["warmups"]
        report["samples"] = []
        for index in range(plan["samples"]):
            sample = copy.deepcopy(templates[route]["samples"][0])
            sample["index"] = index
            report["samples"].append(sample)

        name = f"c{spec['cycle']}-r{spec['repeat']}-{route}.json"
        output = captures / name
        write(output, report)
        stderr = captures / (name + ".stderr")
        stderr.write_bytes(b"")
        runs.append({
            "cycle": spec["cycle"],
            "repeat": spec["repeat"],
            "route": route,
            "command": expected_command(plan, build, case, route),
            "exit_code": 0,
            "seconds": 0.0,
            "output": name,
            "sha256": sha(output),
            "stderr": stderr.name,
            "stderr_sha256": sha(stderr),
        })

    write(captures / "manifest.json", {
        "status": "complete",
        "mode": "synthetic-schema-only",
        "freeze_sha256": sha(packet / "freeze.json"),
        "preflight_sha256": sha(packet / "preflight.json"),
        "runs": runs,
    })
    assert len(runs) == 36


def synthetic_preflight_receipt(packet: Path) -> None:
    """Install the temporary receipt consumed by both validators."""
    write(packet / "preflight.json", {
        "status": "passed",
        "kind": "synthetic-schema-only",
        "freeze_sha256": sha(packet / "freeze.json"),
        "scripts": {
            "preflight.py": sha(packet / "preflight.py"),
            "analyze.py": sha(packet / "analyze.py"),
            "audit.py": sha(packet / "audit.py"),
        },
        "process_count": 36,
    })


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_validators(packet: Path) -> None:
    analyzer = load_module("change0732_preflight_analyzer", packet / "analyze.py")
    audit = load_module("change0732_preflight_audit", packet / "audit.py")

    analyzer.P = packet
    analyzer.ROOT = ROOT
    analyzer.CAPTURES = packet / "captures"
    with contextlib.redirect_stdout(io.StringIO()):
        analyzer.main()

    # Support both the current PACKET spelling and the in-flight P alias while
    # keeping the audit independent from the analyzer implementation.
    audit.P = packet
    audit.PACKET = packet
    audit.ROOT = ROOT
    audit.CAPTURES = packet / "captures"
    with contextlib.redirect_stdout(io.StringIO()):
        audit.main()

    analysis = read(packet / "analysis.json")
    assert analysis.get("status") == "passed"
    assert analysis.get("process_count") == 36


def main() -> None:
    assert (P / "freeze.json").is_file(), "freeze.json is required; run qualify then freeze"
    original_freeze_sha = sha(P / "freeze.json")
    with tempfile.TemporaryDirectory(prefix="litchi-0732-preflight-") as temporary:
        packet = Path(temporary)
        copy_packet(packet)
        rebase_packet_freeze(packet)
        rebase_quality_manifest(packet)
        synthetic_preflight_receipt(packet)
        synthetic_captures(packet)
        run_validators(packet)

    analyzer_path = P / "analyze.py"
    audit_path = P / "audit.py"
    script_path = Path(__file__).resolve()
    receipt = {
        "status": "passed",
        "kind": "synthetic-schema-only",
        "freeze_sha256": original_freeze_sha,
        "scripts": {
            "preflight.py": sha(script_path),
            "analyze.py": sha(analyzer_path),
            "audit.py": sha(audit_path),
        },
        "script_sha256": sha(script_path),
        "analyzer_sha256": sha(analyzer_path),
        "audit_sha256": sha(audit_path),
        "auditor_sha256": sha(audit_path),
        "process_count": 36,
        "synthetic_capture_count": 36,
    }
    write(P / "preflight.json", receipt)
    print("PASS synthetic 36-process analyzer/audit schema integration; no measurement evidence")


if __name__ == "__main__":
    main()
