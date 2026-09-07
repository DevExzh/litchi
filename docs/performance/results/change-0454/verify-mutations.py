#!/usr/bin/env python3
"""Exercise the final-readiness rejection paths in a disposable bundle copy.

The live evidence bundle is never edited.  The script is intentionally small:
it first checks for a released final bundle, then mutates one custody field at
a time and requires ``verify.py --precleanup`` to reject each copy.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
FINAL_GATE_TAGS = (
    "final-strict",
    "final-harness-strict",
    "final-pptx",
    "final-opc",
    "final-harness",
    "final-doc",
    "final-workspace",
    "final-format",
    "final-boundaries",
    "fuzz-lock",
    "fuzz-build",
    "fuzz-smoke",
    "final-fuzz-strict",
    "final-fuzz-format",
)


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def latest_gate(bundle: Path, tag: str) -> Path | None:
    checks = bundle / "checks"
    candidates: list[tuple[int, Path]] = []
    exact = checks / f"{tag}.json"
    if exact.is_file():
        candidates.append((0, exact))
    pattern = re.compile(re.escape(tag) + r"-r([0-9]+)\.json")
    for path in checks.glob(f"{tag}-r*.json"):
        match = pattern.fullmatch(path.name)
        if match is not None:
            candidates.append((int(match.group(1)), path))
    if not candidates:
        return None
    return max(candidates, key=lambda item: (item[0], item[1].name))[1]


def latest_receipt(bundle: Path, tag: str) -> Path | None:
    """Select the newest tagged receipt, including a final exact-name fallback."""
    return latest_gate(bundle, tag)


def ready() -> bool:
    required = (
        ROOT / "candidate-build.json",
        ROOT / "machine.json",
        ROOT / "measurements.json",
        ROOT / "measurements.md",
        ROOT / "final-native-inventory.json",
        ROOT / "final-outcome-comparison.json",
        ROOT / "final-probe-driver.json",
        ROOT / "fuzz-source-amendment.json",
        ROOT / "derivation-amendment.json",
        ROOT / "derive-final.py",
        ROOT / "provider-runs",
        ROOT / "external-runs",
        ROOT / "external-pilots",
    )
    if not all(path.exists() for path in required):
        return False
    for tag in ("final-candidate-build", "final-native-inventory", *FINAL_GATE_TAGS):
        path = latest_gate(ROOT, tag)
        if path is None or load(path).get("status") != "pass":
            return False
    return True


def copy_bundle(parent: Path, name: str) -> Path:
    target = parent / name
    shutil.copytree(ROOT, target, ignore=shutil.ignore_patterns("__pycache__"))
    return target


def run(bundle: Path) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    return subprocess.run(
        [sys.executable, "-B", str(bundle / "verify.py"), "--precleanup"],
        cwd=bundle,
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def expect_rejection(name: str, mutate: Callable[[Path], None], parent: Path) -> str:
    bundle = copy_bundle(parent, name)
    mutate(bundle)
    result = run(bundle)
    if result.returncode == 0:
        raise AssertionError(f"{name}: verifier unexpectedly accepted mutation")
    return name


def final_receipt(bundle: Path) -> Path:
    path = latest_receipt(bundle, "final-native-inventory")
    if path is None:
        raise AssertionError("final-native-inventory receipt is missing from released bundle")
    return path


def mutate_missing_receipt(bundle: Path) -> None:
    checks = bundle / "checks"
    pattern = re.compile(r"final-native-inventory(?:-r[0-9]+)?\.json")
    removed = [path for path in checks.glob("final-native-inventory*.json") if pattern.fullmatch(path.name)]
    if not removed:
        raise AssertionError("final-native-inventory receipts are missing from released bundle")
    for path in removed:
        path.unlink()


def mutate_running_receipt(bundle: Path) -> None:
    path = final_receipt(bundle)
    value = load(path)
    value["status"] = "running"
    for field in ("source_after", "source_unchanged", "finished_utc", "log"):
        value.pop(field, None)
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def mutate_log(bundle: Path) -> None:
    value = load(final_receipt(bundle))
    log = value["log"]
    path = bundle / log["path"]
    raw = path.read_bytes()
    path.write_bytes((b"x" if not raw or raw[:1] != b"x" else b"y") + raw[1:])


def provider_receipt(bundle: Path) -> tuple[Path, dict[str, Any]]:
    path = bundle / "provider-runs" / "0" / "receipt.json"
    return path, load(path)


def mutate_incomplete_lane(bundle: Path) -> None:
    provider_receipt(bundle)[0].unlink()


def rewrite_report(bundle: Path, update: Callable[[dict[str, Any]], None]) -> None:
    receipt_path, receipt = provider_receipt(bundle)
    report_ref = receipt["artifacts"]["report"]
    report_path = bundle / report_ref["path"]
    report = load(report_path)
    update(report)
    raw = (json.dumps(report, indent=2) + "\n").encode()
    report_path.write_bytes(raw)
    report_ref["bytes"] = len(raw)
    report_ref["sha256"] = sha(raw)
    receipt_path.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")


def mutate_report_identity(bundle: Path) -> None:
    rewrite_report(bundle, lambda report: report.__setitem__("binary_sha256", "0" * 64))


def mutate_report_oracle(bundle: Path) -> None:
    def update(report: dict[str, Any]) -> None:
        report["provider_config"]["max_range_bytes"] = 0

    rewrite_report(bundle, update)


def final_gate(bundle: Path) -> Path:
    path = latest_gate(bundle, "final-strict")
    if path is None:
        raise AssertionError("final-strict gate receipt is missing from released bundle")
    return path


def mutate_missing_gate(bundle: Path) -> None:
    checks = bundle / "checks"
    pattern = re.compile(r"final-strict(?:-r[0-9]+)?\.json")
    removed = [path for path in checks.glob("final-strict*.json") if pattern.fullmatch(path.name)]
    if not removed:
        raise AssertionError("final-strict gate receipts are missing from released bundle")
    for path in removed:
        path.unlink()


def mutate_failed_gate(bundle: Path) -> None:
    path = final_gate(bundle)
    value = load(path)
    value["status"] = "failed"
    value["exit_code"] = 1
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def mutate_stale_gate(bundle: Path) -> None:
    path = final_gate(bundle)
    value = load(path)
    current = value["source_before"]["sha256"]
    stale_path = next(
        candidate
        for candidate in sorted((bundle / "sources").glob("*.json"))
        if candidate.stem != current
    )
    raw = stale_path.read_bytes()
    stale = {"path": str(stale_path.relative_to(bundle)), "sha256": sha(raw), "files": len(load(stale_path))}
    value["source_before"] = stale
    value["source_after"] = stale
    value["source_unchanged"] = True
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def mutate_nonfuzz_source_amendment(bundle: Path) -> None:
    path = bundle / "fuzz-source-amendment.json"
    value = load(path)
    value["changed_paths"] = [
        value["allowed_changed_path"],
        "crates/litchi-opc/src/lib.rs",
    ]
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def mutate_renderer_outside_allowed_block(bundle: Path) -> None:
    driver_path = bundle / "derive-final.py"
    mutated = driver_path.read_bytes() + b"\n# amendment mutation outside the permitted renderer block\n"
    driver_path.write_bytes(mutated)
    amendment_path = bundle / "derivation-amendment.json"
    amendment = load(amendment_path)
    amendment["corrected_sha256"] = sha(mutated)
    amendment_path.write_text(json.dumps(amendment, indent=2) + "\n", encoding="utf-8")


def main() -> int:
    if not ready():
        print(json.dumps({"status": "skip", "reason": "final evidence bundle is not released yet"}, sort_keys=True))
        return 0
    cases: list[str] = []
    with tempfile.TemporaryDirectory(prefix="change0454-verify-mutations-") as directory:
        parent = Path(directory)
        mutations = (
            ("missing-final-receipt", mutate_missing_receipt),
            ("running-final-receipt", mutate_running_receipt),
            ("changed-final-log", mutate_log),
            ("incomplete-provider-lane", mutate_incomplete_lane),
            ("report-identity-mismatch", mutate_report_identity),
            ("report-oracle-mismatch", mutate_report_oracle),
            ("missing-final-gate", mutate_missing_gate),
            ("failed-final-gate", mutate_failed_gate),
            ("stale-final-gate-source", mutate_stale_gate),
            ("non-fuzz-source-amendment", mutate_nonfuzz_source_amendment),
            ("renderer-outside-allowed-block", mutate_renderer_outside_allowed_block),
        )
        for name, mutate in mutations:
            cases.append(expect_rejection(name, mutate, parent))
    print(json.dumps({"status": "pass", "rejected": cases}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
