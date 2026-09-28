"""Audit the retained 0799 build and quality failures without rerunning them."""

import argparse
import re
from pathlib import Path

import custody as c


p = c.P

HELPER_PATHS = [
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
]


def artifact(path: Path) -> dict:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {
        "path": str(path.relative_to(p.resolve())),
        "bytes": path.stat().st_size,
        "sha256": c.sha(path),
    }


def external_descriptor(
    value: dict, label: str, *, allow_replaced_path: bool = False
) -> dict:
    path = Path(value["path"])
    assert path.is_absolute(), label
    size = value["bytes"]
    digest = value["sha256"]
    assert isinstance(size, int) and size > 0, label
    assert re.fullmatch(r"[0-9a-f]{64}", digest), label
    if path.is_file() and not path.is_symlink():
        if path.stat().st_size != size or c.sha(path) != digest:
            # The failed run's original output path is reused by the later
            # successful build. Its exact failed binary is checked below at
            # retained_binary, while cleanup records both identities.
            assert allow_replaced_path, label
    else:
        cleanup_path = p / "cleanup.json"
        assert cleanup_path.is_file(), (label, path)
        cleanup = c.read(cleanup_path)
        expected = {"path": str(path), "bytes": size, "sha256": digest}
        if label == "original binary":
            # The later successful build reused this path, so cleanup's
            # removed_binaries identity is the final binary, not this failed
            # run's original bytes. The path itself must still be witnessed.
            assert any(item["path"] == str(path) for item in cleanup["removed_binaries"])
        else:
            assert expected in cleanup["removed_failed_binaries"]
        assert cleanup["target_removed"] is True
        assert cleanup["target"] == str(c.TARGET)
        assert not c.TARGET.exists()
    return {"path": str(path), "bytes": size, "sha256": digest}


def descriptor_matches(path: Path, expected: dict) -> dict:
    actual = artifact(path)
    assert actual["bytes"] == expected["bytes"]
    assert actual["sha256"] == expected["sha256"]
    return actual


def validate_hash_map(root: Path, expected: dict[str, str]) -> list[dict]:
    root = root.resolve()
    actual_paths = {
        str(path.relative_to(root))
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    assert actual_paths == set(expected), (root, actual_paths ^ set(expected))
    files = []
    for name, digest in sorted(expected.items()):
        path = root / name
        current = artifact(path)
        assert current["sha256"] == digest, name
        files.append(current)
    return files


def retained_log(row: dict, root: Path, markers: list[str]) -> dict:
    metadata = row["log"]
    original = Path(metadata["path"])
    assert original.is_absolute()
    retained = root / original.name
    current = descriptor_matches(retained, metadata)
    text = retained.read_text()
    for marker in markers:
        assert marker in text, (retained, marker)
    return {
        "original": metadata["path"],
        "retained": current,
        "markers": markers,
    }


def receipt_audit(
    root: Path,
    receipt_name: str,
    expected_exit_codes: list[int],
    markers: dict[int, list[str]],
) -> dict:
    rows = c.read(root / receipt_name)
    assert len(rows) == len(expected_exit_codes)
    logs = []
    previous_end = None
    for index, (row, expected) in enumerate(zip(rows, expected_exit_codes)):
        assert row["exit_code"] == expected
        assert row["started"] <= row["ended"]
        if previous_end is not None:
            assert previous_end <= row["started"]
        previous_end = row["ended"]
        logs.append(retained_log(row, root, markers.get(index, [])))
    return {
        "commands": [row["command"] for row in rows],
        "exit_codes": expected_exit_codes,
        "logs": logs,
    }


def patch_artifact(name: str, normalized: bool) -> dict:
    path = p / name
    text = path.read_text()
    old = [line.split()[1] for line in text.splitlines() if line.startswith("--- ")]
    new = [line.split()[1] for line in text.splitlines() if line.startswith("+++ ")]
    assert len(old) == len(new) == len(HELPER_PATHS)
    if normalized:
        assert old == [f"a/{name}" for name in HELPER_PATHS]
        assert new == [f"b/{name}" for name in HELPER_PATHS]
        assert "candidate/before" not in text
        assert "candidate/after" not in text
    else:
        assert all(value.startswith("candidate/before/") for value in old)
        assert all(value.startswith("candidate/after/") for value in new)
    return artifact(path)


def audit_quality_failure(number: int) -> dict:
    root = p / f"quality-failed-{number}"
    relocation = c.read(root / "relocation.json")
    assert relocation["quality"] == root.name
    assert relocation["candidate"] == f"candidate-failed-{number}"
    assert relocation["test-src"] == f"test-src-failed-{number}"
    if number == 0:
        expected = [0, 0, 0, 0, 101]
        markers = {
            4: [
                "error[E0583]",
                "file not found for module `tests`",
                "litchi-ole-common",
            ]
        }
    else:
        expected = [0, 0, 0, 0, 0, 101]
        markers = {
            5: [
                "clippy::collapsible-if",
                "this `if` statement can be collapsed",
            ]
        }
    result = receipt_audit(root, "receipts.json", expected, markers)

    archive_root = p / relocation["candidate"]
    archive_inputs = c.read(root / "archive-inputs.json")
    archive = validate_hash_map(archive_root, archive_inputs)
    test_root = p / relocation["test-src"]
    mirror = validate_hash_map(test_root, relocation["inputs"])
    patch = patch_artifact(f"candidate-failed-{number}.patch", normalized=True)
    result.update(
        {
            "id": root.name,
            "stage": "isolated helper quality",
            "reason": relocation["fix"],
            "relocation": artifact(root / "relocation.json"),
            "archive_inputs": artifact(root / "archive-inputs.json"),
            "candidate_archive": archive,
            "test_source_mirror": mirror,
            "normalized_patch": patch,
        }
    )
    return result


def audit_build_failure() -> dict:
    root = p / "build-failed-0"
    relocation = c.read(root / "relocation.json")
    assert relocation["build"] == root.name
    assert relocation["probe-src"] == "probe-src-failed-0"
    assert "40cases" in relocation["reason"]
    assert "duplicate-unquoted-after-2" in relocation["reason"]

    inputs = c.read(root / "inputs.json")
    frozen = []
    for name, digest in sorted(inputs["frozen"].items()):
        current = artifact(p / name)
        assert current["sha256"] == digest, name
        frozen.append(current)
    probe = validate_hash_map(p / relocation["probe-src"], inputs["probe"])
    assert inputs["probe"] == relocation["probe_inputs"]

    result = receipt_audit(
        root,
        "commands.json",
        [0, 0, 0, 0, 0, 0],
        {},
    )
    rows = c.read(root / "commands.json")
    assert len(rows) == 6
    for index in range(4):
        assert rows[index]["command"][0] in {"rustfmt", "cargo"}
    assert rows[4]["command"][-1] == "--list-cases"
    assert rows[5]["command"][-1] == "--self-check"

    cases = root / "cases.json"
    self_check = root / "self-check.json"
    for index, output in ((4, cases), (5, self_check)):
        assert rows[index]["output"]["path"].endswith(output.name)
        descriptor_matches(output, rows[index]["output"])
    failed_cases = c.read(cases)
    failed_self_check = c.read(self_check)
    assert len(failed_cases) == len(failed_self_check["cases"]) == 40
    case_ids = [case["id"] for case in failed_cases]
    self_ids = [case["id"] for case in failed_self_check["cases"]]
    assert len(set(case_ids)) == len(case_ids) == 40
    assert case_ids == self_ids
    assert failed_self_check["all_checks_passed"] is True
    assert case_ids.count("duplicate-unquoted-after-2") == 1

    plan = c.read(p / "plan.json")
    fixtures = c.read(p / "fixtures.json")
    assert plan["case_count"] == 39
    assert len(fixtures) == 39
    assert "duplicate-unquoted-after-2" not in {fixture["id"] for fixture in fixtures}

    original_binary = external_descriptor(
        relocation["original_binary"],
        "original binary",
        allow_replaced_path=True,
    )
    retained_binary = external_descriptor(relocation["retained_binary"], "retained failed binary")
    result.update(
        {
            "id": root.name,
            "stage": "control probe build and catalog",
            "reason": relocation["reason"],
            "relocation": artifact(root / "relocation.json"),
            "inputs": artifact(root / "inputs.json"),
            "frozen_inputs": frozen,
            "probe_source": probe,
            "retained_cases": artifact(cases),
            "retained_self_check": artifact(self_check),
            "original_binary": original_binary,
            "retained_failed_binary": retained_binary,
        }
    )
    return result


def result() -> dict:
    return {
        "schema": "litchi.performance.0799.failure-audit.v1",
        "production_adoption": False,
        "production_source_changed": False,
        "failed_attempts": [
            audit_build_failure(),
            audit_quality_failure(0),
            audit_quality_failure(1),
        ],
        "patches": {
            "initial_unapplyable": patch_artifact(
                "candidate-archive-paths.patch", normalized=False
            ),
            "quality_failed_0": patch_artifact(
                "candidate-failed-0.patch", normalized=True
            ),
            "quality_failed_1": patch_artifact(
                "candidate-failed-1.patch", normalized=True
            ),
        },
        "path_policy": "receipt paths are resolved through each retained failure archive; no failed command is rerun",
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
    print("0799 failure audit PASS")
