"""Audit retained 0800 failure and edge evidence without rerunning Cargo."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import custody as c


p = c.P

HELPER_PATHS = [
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
]


def archive_name(source_name: str) -> str:
    if source_name.endswith("crates/litchi-opc/src/xml_attributes/tests.rs"):
        return "litchi-opc-xml_attributes-tests.rs"
    return Path(source_name).parent.parent.name + "-xml_attributes.rs"


def packet_path(value: str | Path) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path.resolve()
    for candidate in (p / path, c.ROOT / path):
        if candidate.exists():
            return candidate.resolve()
    return (p / path).resolve()


def relative(path: Path) -> str:
    return str(path.resolve().relative_to(p.resolve()))


def artifact(path: Path) -> dict:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {"path": relative(path), "bytes": path.stat().st_size, "sha256": c.sha(path)}


def descriptor_matches(path: Path, expected: dict) -> dict:
    actual = artifact(path)
    assert actual["bytes"] == expected["bytes"]
    assert actual["sha256"] == expected["sha256"]
    return actual


def external_descriptor(value: dict, label: str) -> dict:
    path = Path(value["path"])
    assert path.is_absolute(), label
    assert isinstance(value["bytes"], int) and value["bytes"] > 0, label
    assert re.fullmatch(r"[0-9a-f]{64}", value["sha256"]), label
    if path.is_file() and not path.is_symlink():
        assert path.stat().st_size == value["bytes"], label
        assert c.sha(path) == value["sha256"], label
    else:
        cleanup_path = p / "cleanup.json"
        assert cleanup_path.is_file(), (label, path)
        cleanup = c.read(cleanup_path)
        assert cleanup["target"] == str(c.TARGET)
        assert cleanup["target_removed"] is True
        assert not c.TARGET.exists()
        assert value in cleanup["removed_binaries"], (label, value)
    return {
        "path": value["path"],
        "bytes": value["bytes"],
        "sha256": value["sha256"],
    }


def validate_tree(root: Path, expected: dict[str, str]) -> list[dict]:
    root = root.resolve()
    actual_names = {
        str(path.relative_to(root))
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    assert actual_names == set(expected), (root, actual_names ^ set(expected))
    result = []
    for name, digest in sorted(expected.items()):
        current = artifact(root / name)
        assert current["sha256"] == digest, name
        result.append(current)
    return result


def validate_packet_inputs(inputs: dict[str, str], *, archive: Path | None = None) -> list[dict]:
    """Validate packet hashes, mapping old correction paths into an archive."""
    result = []
    for name, digest in sorted(inputs.items()):
        if archive is not None and name == "correction.patch":
            path = archive.parent / "correction-failed-0.patch"
        elif archive is not None and name.startswith("correction/"):
            archived_name = name.removeprefix("correction/")
            path = archive / archived_name
        else:
            path = packet_path(name)
        current = artifact(path)
        assert current["sha256"] == digest, name
        result.append(current)
    return result


def validate_patch(path: Path) -> dict:
    text = path.read_text()
    old = {line.split()[1] for line in text.splitlines() if line.startswith("--- ")}
    new = {line.split()[1] for line in text.splitlines() if line.startswith("+++ ")}
    assert old == {f"a/{name}" for name in HELPER_PATHS}
    assert new == {f"b/{name}" for name in HELPER_PATHS}
    return artifact(path)


def retained_log(root: Path, row: dict, markers: list[str]) -> dict:
    metadata = row["log"]
    original = Path(metadata["path"])
    assert original.is_absolute()
    retained = root / original.name
    current = descriptor_matches(retained, metadata)
    text = retained.read_text(errors="replace")
    for marker in markers:
        assert marker in text, (retained, marker)
    return {"original": metadata["path"], "retained": current, "markers": markers}


def receipt_audit(
    root: Path,
    expected_exit_codes: list[int],
    markers: dict[int, list[str]],
) -> dict:
    rows = c.read(root / "receipts.json")
    assert len(rows) == len(expected_exit_codes)
    previous_end = None
    logs = []
    for index, (row, expected) in enumerate(zip(rows, expected_exit_codes)):
        assert row["exit_code"] == expected
        assert row["started"] <= row["ended"]
        if previous_end is not None:
            assert previous_end <= row["started"]
        previous_end = row["ended"]
        logs.append(retained_log(root, row, markers.get(index, [])))
    return {
        "commands": [row["command"] for row in rows],
        "exit_codes": expected_exit_codes,
        "logs": logs,
    }


def audit_correction_archive() -> dict:
    manifest_path = p / "correction/manifest.json"
    manifest = c.read(manifest_path)
    assert manifest["schema"] == "litchi.performance.0800.correctness.v1"
    assert manifest["base"] == c.read(p / "origin.json")["base"]
    assert manifest["performance_claim"] is False
    assert set(manifest["changed"]) == set(HELPER_PATHS)
    baseline = c.read(p / "source.json")["files"]
    after_hashes = {}
    before_hashes = {}
    tree_hashes = {"manifest.json": c.sha(manifest_path)}
    for name, parts in manifest["changed"].items():
        before = packet_path(parts["before"]["path"]) if parts["before"] else None
        after = packet_path(parts["after"]["path"]) if parts["after"] else None
        assert before is not None and after is not None
        before_hashes[f"before/{before.name}"] = parts["before"]["sha256"]
        after_hashes[f"after/{after.name}"] = parts["after"]["sha256"]
        assert baseline[name] == c.sha(before)
        assert c.sha(after) == parts["after"]["sha256"]
        assert c.sha(c.ROOT / name) == parts["after"]["sha256"]
        tree_hashes[f"before/{before.name}"] = parts["before"]["sha256"]
        tree_hashes[f"after/{after.name}"] = parts["after"]["sha256"]
    archive = validate_tree(p / "correction", tree_hashes)
    text = (p / "correction/after/litchi-opc-xml_attributes.rs").read_text()
    assert "consume the first byte" in text
    test_text = (p / "correction/after/litchi-opc-xml_attributes-tests.rs").read_text()
    assert "=long_name" in test_text and "==n0" not in test_text
    return {
        "manifest": artifact(manifest_path),
        "patch": validate_patch(p / "correction.patch"),
        "archive": archive,
        "before_hashes": before_hashes,
        "after_hashes": after_hashes,
    }


def audit_failed_quality() -> dict:
    root = p / "quality-failed-0"
    relocation = c.read(root / "relocation.json")
    assert relocation["quality"] == root.name
    assert relocation["correction"] == "correction-failed-0"
    assert relocation["correction.patch"] == "correction-failed-0.patch"
    inputs = c.read(root / "inputs.json")
    correction_inputs = {
        name: digest for name, digest in inputs.items() if name.startswith("correction/")
    }
    archive = validate_tree(
        p / relocation["correction"],
        {name.removeprefix("correction/"): digest for name, digest in correction_inputs.items()},
    )
    packet_inputs = validate_packet_inputs(inputs, archive=p / relocation["correction"])
    failed_source = c.read(root / "source.json")
    baseline = c.read(p / "source.json")["files"]
    assert failed_source["revision"] == c.read(p / "source.json")["revision"]
    for name, digest in failed_source["files"].items():
        if name in HELPER_PATHS:
            archive_key = "correction/after/" + archive_name(name)
            assert digest == correction_inputs[archive_key], name
        else:
            assert digest == baseline[name], name
    result = receipt_audit(
        root,
        [0, 0, 101],
        {
            2: [
                "unusual_duplicate_keys_keep_error_precedence_at_the_switch",
                "assertion failed",
                "test result: FAILED",
            ]
        },
    )
    result.update(
        {
            "id": root.name,
            "reason": relocation["reason"],
            "fix": relocation["fix"],
            "relocation": artifact(root / "relocation.json"),
            "inputs": packet_inputs,
            "source": artifact(root / "source.json"),
            "correction_archive": archive,
            "patch": validate_patch(p / relocation["correction.patch"]),
        }
    )
    return result


def expected_edge_lines(controlled: bool) -> list[str]:
    if controlled:
        return [
            "accepted_prefix=1 baseline=Some(Duplicated(10, 2)) quick_xml=Some(Duplicated(10, 2)) equal=true corrected=Some(Duplicated(10, 2))",
            "accepted_prefix=31 baseline=Some(Duplicated(292, 2)) quick_xml=Some(Duplicated(292, 2)) equal=true corrected=Some(Duplicated(292, 2))",
            "accepted_prefix=32 baseline=Some(ExpectedQuote(319, 34)) quick_xml=Some(Duplicated(302, 2)) equal=false corrected=Some(Duplicated(302, 2))",
            "accepted_prefix=33 baseline=Some(ExpectedQuote(329, 34)) quick_xml=Some(Duplicated(312, 2)) equal=false corrected=Some(Duplicated(312, 2))",
            "accepted_prefix=34 baseline=Some(ExpectedQuote(339, 34)) quick_xml=Some(Duplicated(322, 2)) equal=false corrected=Some(Duplicated(322, 2))",
        ]
    return [
        "accepted_prefix=1 baseline=Some(Duplicated(10, 2)) quick_xml=Some(Duplicated(10, 2)) equal=true",
        "accepted_prefix=31 baseline=Some(Duplicated(292, 2)) quick_xml=Some(Duplicated(292, 2)) equal=true",
        "accepted_prefix=32 baseline=Some(ExpectedQuote(319, 34)) quick_xml=Some(Duplicated(302, 2)) equal=false",
        "accepted_prefix=33 baseline=Some(ExpectedQuote(329, 34)) quick_xml=Some(Duplicated(312, 2)) equal=false",
        "accepted_prefix=34 baseline=Some(ExpectedQuote(339, 34)) quick_xml=Some(Duplicated(322, 2)) equal=false",
    ]


def audit_edge_evidence() -> dict:
    edge = p / "edge-probe"
    initial = c.read(edge / "initial-run.json")
    assert initial["exit_code"] == 0
    assert initial["timestamps_recorded"] is False
    assert "--locked" not in initial["command"]
    assert initial["baseline"]["sha256"] == c.sha(edge / "src/baseline.rs")
    assert initial["lock"]["sha256"] == c.sha(edge / "Cargo.lock")
    for output in initial["outputs"]:
        descriptor_matches(packet_path(output["path"]), output)
    initial_lines = (edge / "output.log").read_text().splitlines()
    assert initial_lines == expected_edge_lines(False)

    controlled_lines = [
        line
        for line in (p / "quality/1.log").read_text().splitlines()
        if line.startswith("accepted_prefix=")
    ]
    assert controlled_lines == expected_edge_lines(True)
    corrected = c.sha(edge / "src/corrected.rs")
    assert corrected == c.sha(p / "correction/after/litchi-opc-xml_attributes.rs")
    return {
        "initial_exploratory": {
            "run": artifact(edge / "initial-run.json"),
            "source": artifact(edge / "initial-main.rs.txt"),
            "baseline": artifact(edge / "src/baseline.rs"),
            "corrected": artifact(edge / "src/corrected.rs"),
            "lock": artifact(edge / "Cargo.lock"),
            "build_log": artifact(edge / "build.log"),
            "output": artifact(edge / "output.log"),
        },
        "controlled_locked": {
            "log": artifact(p / "quality/1.log"),
            "rows": 5,
            "mismatches": [32, 33, 34],
            "corrected_matches": [1, 31, 32, 33, 34],
        },
    }


def audit_deferred_candidate() -> dict:
    root = p / "deferred-candidate"
    manifest_path = root / "manifest.json"
    manifest = c.read(manifest_path)
    assert manifest["schema"] == "litchi.performance.0800.candidate-manifest.v1"
    assert manifest["base_commit"] == c.read(p / "origin.json")["base"]
    assert manifest["production_adoption"] is False
    assert manifest["production_source_changed"] is False
    assert manifest["candidate_name"] == "checked-short-prefix-no-replay"
    assert manifest["status"] == "untested_deferred_requires_rebase"
    assert "no build or measurement approval" in manifest["deferred_reason"]
    assert len(manifest["before"]) == len(manifest["after"]) == 6
    baseline = c.read(p / "source.json")["files"]
    before_expected = {}
    for name in HELPER_PATHS:
        before_expected[f"before/{archive_name(name)}"] = baseline[name]
    assert set(before_expected) == {
        str(path.relative_to(root))
        for path in (root / "before").iterdir()
        if path.is_file()
    }
    archive_files = []
    for name, digest in sorted(before_expected.items()):
        current = artifact(root / name)
        assert current["sha256"] == digest, name
        archive_files.append(current)
    after_names = {
        str(path.relative_to(root))
        for path in (root / "after").iterdir()
        if path.is_file()
    }
    assert after_names == {f"after/{archive_name(name)}" for name in HELPER_PATHS}
    for path in sorted((root / "after").iterdir()):
        assert path.is_file()
        current = artifact(path)
        source_name = next(
            name for name in HELPER_PATHS if archive_name(name) == path.name
        )
        assert current["sha256"] != c.sha(c.ROOT / source_name)
        archive_files.append(current)
    archive_files.extend([artifact(root / "design.md"), artifact(manifest_path)])
    design_text = (root / "design.md").read_text().lower()
    assert "source-only" in design_text
    assert "no production edit" in design_text
    assert "build receipt" in design_text
    return {
        "manifest": artifact(manifest_path),
        "design": artifact(root / "design.md"),
        "patch": validate_patch(p / "deferred-candidate.patch"),
        "archive": archive_files,
        "status": "source-only, untested, not adopted",
    }


def audit_final_quality() -> dict:
    complete = c.read(p / "quality/complete.json")
    receipts = c.read(p / "quality/receipts.json")
    assert len(receipts) == len(complete["rows"]) == 4
    assert all(row["exit_code"] == 0 for row in receipts)
    assert complete["scope"].endswith("no performance claim")
    for row in receipts:
        assert row["started"] <= row["ended"]
        descriptor_matches(packet_path(row["log"]["path"]), row["log"])
    summary = c.read(p / "test-summary.json")
    assert summary == {
        "failed": 0,
        "filtered": 0,
        "ignored": 3,
        "measured": 0,
        "passed": 1538,
        "scope": "Full default-feature tests for five affected production crates, including integration and doctest suites",
        "source": summary["source"],
        "suites": 67,
    }
    descriptor_matches(packet_path(summary["source"]["path"]), summary["source"])
    return {
        "complete": artifact(p / "quality/complete.json"),
        "receipts": artifact(p / "quality/receipts.json"),
        "test_summary": artifact(p / "test-summary.json"),
        "exit_codes": [row["exit_code"] for row in receipts],
    }


def result() -> dict:
    identities = c.read(p / "binary-identities.json")
    assert identities["order"] == ["initial_exploratory", "controlled_locked"]
    assert len(identities["binaries"]) == 2
    binaries = [external_descriptor(item, f"binary {index}") for index, item in enumerate(identities["binaries"])]
    return {
        "schema": "litchi.performance.0800.failure-audit.v1",
        "production_correction": True,
        "performance_claim": False,
        "binary_identities": {
            "artifact": artifact(p / "binary-identities.json"),
            "order": identities["order"],
            "binaries": binaries,
        },
        "correction": audit_correction_archive(),
        "failed_attempts": [audit_failed_quality()],
        "edge_evidence": audit_edge_evidence(),
        "deferred_candidate": audit_deferred_candidate(),
        "final_quality": audit_final_quality(),
        "path_policy": "old correction input paths are resolved by archive-relative suffix; retained receipt logs are resolved by basename; no command is rerun",
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output = p / "failure-audit.json"
    expected = result()
    if args.write:
        assert not output.exists(), output
        c.write(output, expected)
    else:
        assert c.read(output) == expected
    print("0800 failure audit PASS")
