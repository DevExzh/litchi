"""Verify and remove the owned 0808 target after the pre-measurement stop.

This is deliberately separate from the full six-binary cleanup. It retains
the candidate archive, qualification reports, probe receipts, and failed
after-quality log while removing only the owned target and its three before
binaries. It must run after root has restored the baseline source.
"""

from pathlib import Path
import shutil

import custody as c


P = c.P
TARGET = c.TARGET
OUT = P / "early-stop-cleanup.json"


def require(value: bool, message: str) -> None:
    if not value:
        raise RuntimeError(message)


def packet_artifact(path: Path, label: str) -> dict[str, object]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing")
    return c.artifact(path)


def target_binary(value: dict[str, object], label: str) -> dict[str, object]:
    path = Path(value["path"])
    require(path.is_absolute() and path.parent == TARGET, f"{label}: path outside owned target")
    require(path.is_file() and not path.is_symlink(), f"{label}: binary missing")
    require(path.stat().st_size == value["bytes"], f"{label}: binary size changed")
    require(c.sha(path) == value["sha256"], f"{label}: binary hash changed")
    return {
        "path": str(path),
        "bytes": value["bytes"],
        "sha256": value["sha256"],
    }


def main() -> None:
    require(not OUT.exists(), "refusing to overwrite early-stop cleanup witness")
    plan = c.read(P / "early-stop-plan.json")
    require(plan["schema"] == "litchi.performance.0808.early-stop-plan.v1",
            "early-stop plan schema changed")
    require(plan["status"] == "stopped-before-after-build",
            "this driver is only for the terminal early-stop path")

    # Root-owned restoration is a precondition. This prevents cleanup from
    # hiding a candidate source tree or deleting evidence before its decision.
    before_source = c.read(P / "build-before/source.json")
    restored_source = c.read(P / "restored-source.json")
    disposition = c.read(P / "disposition.json")
    require(disposition["status"] == "rejected"
            and disposition["production_change_retained"] is False,
            "baseline restoration disposition is not recorded")
    require(restored_source == before_source, "restored source differs from before source")
    require(c.source() == before_source, "current production source is not restored")

    # Preserve the reviewed candidate and the application witness. These are
    # checked before any target removal and are never deleted by this driver.
    for path in (
        P / "candidate/manifest.json",
        P / "candidate/candidate.patch",
        P / "candidate/before/codec.rs",
        P / "candidate/after/codec.rs",
        P / "candidate-review.md",
        P / "application.json",
    ):
        packet_artifact(path, f"retained evidence {path.relative_to(P)}")
    application = c.read(P / "application.json")
    after_quality_source = c.read(P / "quality-after/source.json")
    require(after_quality_source == application["source"],
            "after quality source does not match the applied candidate witness")

    qualification_complete = c.read(P / "qualification/complete.json")
    require(qualification_complete["children"] == 18
            and qualification_complete["reports"] == 18
            and qualification_complete["samples"] == 18,
            "qualification cardinality is not 18 reports / 18 samples")
    qualification_audit_path = P / "qualification-audit.json"
    qualification_audit = c.read(qualification_audit_path)
    require(
        qualification_audit["schema"] == "litchi.performance.0808.qualification-audit.v1"
        and qualification_audit["passed"] is True
        and qualification_audit["accepted_before_application"] is True
        and qualification_audit["reports"] == 18
        and qualification_audit["samples"] == 18,
        "independent qualification acceptance receipt is incomplete",
    )
    qualification_receipts_path = P / "qualification/receipts.json"
    qualification_receipts = c.read(qualification_receipts_path)
    require(len(qualification_receipts) == 18, "qualification receipt count changed")
    for index, receipt in enumerate(qualification_receipts):
        require(
            receipt["exit_code"] == 0
            and receipt["lane"] == "qualification"
            and receipt["leg"] == "before"
            and receipt["block"] == 0
            and receipt["report"]["path"].endswith(".json"),
            f"qualification receipt {index} failed",
        )
        report = c.read(Path(receipt["report"]["path"]))
        require(
            report["samples_requested"] == 1
            and report["warmup"] == 0
            and len(report["samples"]) == 1,
            f"qualification report {index} cardinality changed",
        )

    probe_complete = c.read(P / "probe-quality-before/complete.json")
    probe_receipts_path = P / "probe-quality-before/receipts.json"
    probe_receipts = c.read(probe_receipts_path)
    require(
        probe_complete["gate_count"] == 3
        and probe_complete["tests_passed"] == 36
        and len(probe_receipts) == 3
        and all(row["exit_code"] == 0 for row in probe_receipts),
        "before probe quality receipt is incomplete",
    )

    # The production after run is intentionally a four-row terminal receipt:
    # fmt, check, and test pass; warning-denied all-targets Clippy fails.
    quality_checks_path = P / "quality-after/checks.json"
    quality_checks = c.read(quality_checks_path)
    require(len(quality_checks) == 4, "after quality continued past the first failure")
    require([row["exit_code"] for row in quality_checks] == [0, 0, 0, 101],
            "after quality pass/fail sequence changed")
    failed = quality_checks[3]
    require(
        failed["command"][:2] == ["cargo", "clippy"]
        and "--all-targets" in failed["command"]
        and "-D" in failed["command"]
        and "warnings" in failed["command"],
        "failed after gate is not warning-denied all-targets Clippy",
    )
    clippy_log = P / "quality-after/03.log"
    clippy_text = clippy_log.read_text(encoding="utf-8")
    for marker in (
        "clippy::err-expect",
        "crates/litchi-pptx/src/opened/tests.rs:464",
        "crates/litchi-pptx/src/opened/tests.rs:538",
        "crates/litchi-pptx/src/opened/tests.rs:557",
    ):
        require(marker in clippy_text, f"after Clippy diagnostic missing: {marker}")
    require(not (P / "build-after").exists(), "after build unexpectedly exists")
    require(not (P / "native").exists(), "native lane unexpectedly exists")
    require(not (P / "allocation").exists(), "allocation lane unexpectedly exists")
    require(not (P / "profiles").exists(), "profile lane unexpectedly exists")

    build_before = c.read(P / "build-before/build.json")
    binaries = []
    require(set(build_before["binaries"]) == {"native", "allocation", "profile"},
            "before binary set changed")
    require(TARGET.is_dir() and not TARGET.is_symlink(), "owned target is unavailable")
    for kind in ("native", "allocation", "profile"):
        value = build_before["binaries"][kind]
        path = Path(value["path"])
        require(path.name == f"before-{kind}", f"unexpected before binary name: {kind}")
        binaries.append({"kind": kind, **target_binary(value, f"before {kind}")})
    require(len({row["path"] for row in binaries}) == 3,
            "before binary identities are not distinct")

    # Refuse a target containing symlinks before the destructive removal.
    for path in TARGET.rglob("*"):
        require(not path.is_symlink(), f"owned target contains symlink: {path}")
    target_bytes = sum(path.stat().st_size for path in TARGET.rglob("*") if path.is_file())
    shutil.rmtree(TARGET)
    require(not TARGET.exists(), "owned target was not removed")

    manifest = {
        "schema": "litchi.performance.0808.early-stop-cleanup.v1",
        "status": "stopped-before-after-build",
        "reason": "after production quality stopped at warning-denied Clippy gate 4",
        "source_restored": True,
        "source": packet_artifact(P / "restored-source.json", "restored source"),
        "qualification": {
            "complete": packet_artifact(P / "qualification/complete.json",
                                        "qualification complete"),
            "audit": packet_artifact(qualification_audit_path,
                                     "qualification audit"),
            "receipts": packet_artifact(qualification_receipts_path,
                                        "qualification receipts"),
            "reports": 18,
            "samples": 18,
        },
        "probe_quality_before": {
            "complete": packet_artifact(P / "probe-quality-before/complete.json",
                                        "probe quality complete"),
            "receipts": packet_artifact(probe_receipts_path,
                                        "probe quality receipts"),
            "gates": 3,
            "tests_passed": 36,
        },
        "production_quality_after": {
            "checks": packet_artifact(quality_checks_path, "after quality checks"),
            "failed_log": packet_artifact(clippy_log, "after Clippy log"),
            "gates": 4,
            "passes": 3,
            "failures": 1,
            "failed_gate": 4,
            "diagnostic": "clippy::err-expect at opened/tests.rs:464, :538, and :557",
        },
        "target": str(TARGET),
        "target_removed": True,
        "verified_before_removal": True,
        "removed_target_bytes": target_bytes,
        "removed_binaries": binaries,
        "candidate_and_failure_retained": True,
        "workflow_lanes_started": [],
    }
    c.write(OUT, manifest)
    print("0808 early-stop cleanup PASS: removed owned target and three before binaries")


if __name__ == "__main__":
    try:
        main()
    except (AssertionError, KeyError, OSError, RuntimeError) as error:
        print(f"early-stop cleanup failed: {error}")
        raise SystemExit(1)
