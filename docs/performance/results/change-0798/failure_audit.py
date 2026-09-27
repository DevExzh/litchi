"""Audit retained 0798 setup failures and their exact relocations."""

import argparse
from pathlib import Path

import custody as c


p = c.P


def rooted(value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else p / path


def rel(path: Path) -> str:
    return str(path.resolve().relative_to(p.resolve()))


def artifact(path: Path) -> dict:
    path = path.resolve()
    assert path.is_file() and not path.is_symlink(), path
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": c.sha(path)}


def archived_tree(root: Path) -> list[dict]:
    return [artifact(path) for path in sorted(root.rglob("*")) if path.is_file()]


def relocated_path(original: str, before: Path, after: Path) -> Path:
    original_path = rooted(original).resolve()
    before = rooted(before).resolve()
    after = rooted(after).resolve()
    return after / original_path.relative_to(before)


def relocated_log(receipt: dict, before: Path, after: Path) -> tuple[Path, dict]:
    metadata = receipt["log"]
    retained = relocated_path(metadata["path"], before, after)
    actual = artifact(retained)
    assert actual["bytes"] == metadata["bytes"]
    assert actual["sha256"] == metadata["sha256"]
    return retained, {
        "original": metadata["path"],
        "retained": actual,
    }


def receipt_audit(
    root: Path,
    relocation: dict,
    expected_exit_codes: list[int],
    markers: dict[int, list[str]],
    before_key: str,
    after_key: str,
    receipt_name: str = "receipts.json",
) -> dict:
    rows = c.read(root / receipt_name)
    assert len(rows) == len(expected_exit_codes)
    before = Path(relocation[before_key])
    after = Path(relocation[after_key])
    logs = []
    previous_end = None
    for index, (row, expected) in enumerate(zip(rows, expected_exit_codes)):
        assert row["exit_code"] == expected
        assert row["started"] <= row["ended"]
        if previous_end is not None:
            assert previous_end <= row["started"]
        previous_end = row["ended"]
        retained, log = relocated_log(row, before, after)
        text = retained.read_text()
        for marker in markers.get(index, []):
            assert marker in text, (rel(retained), marker)
        log["markers"] = markers.get(index, [])
        logs.append(log)
    return {
        "commands": [row["command"] for row in rows],
        "exit_codes": expected_exit_codes,
        "logs": logs,
    }


def audit_lock_failure() -> dict:
    root = p / "lock-generation-failed-0"
    failure = c.read(root / "failure.json")
    assert failure["exit_code"] is None
    assert failure["returncode_recorded"] is False
    files = []
    for name, expected_hash in sorted(failure["files"].items()):
        path = root / name
        current = artifact(path)
        assert current["sha256"] == expected_hash
        files.append(current)
    error = artifact(root / "error.log")
    assert "does not have that feature" in (root / "error.log").read_text()
    resolved = c.read(p / "lock-generation.json")
    assert resolved["exit_code"] == 0
    assert c.sha(p / "probe-src/Cargo.lock") == resolved["lock"]["sha256"]
    return {
        "id": "lock-generation-failed-0",
        "stage": failure["stage"],
        "reason": failure["reason"],
        "original_returncode": None,
        "returncode_recorded": False,
        "retained_files": files,
        "error_log": error,
        "resolved_lock_generation": {
            "receipt": artifact(p / "lock-generation.json"),
            "lock": artifact(p / "probe-src/Cargo.lock"),
        },
    }


def audit_quality_failure(number: int) -> dict:
    root = p / f"quality-failed-{number}"
    relocation = c.read(root / "relocation.json")
    assert relocation["production_changed"] is False
    if number == 0:
        expected = [0, 101]
        markers = {1: ["error[E0618]", "shadowed by the local binding"]}
    else:
        expected = [0, 0, 101]
        markers = {2: ["clippy::needless-borrow", "change this to: `self.tag`"]}
    result = receipt_audit(
        root,
        relocation,
        expected,
        markers,
        "quality_prefix_before",
        "quality_prefix_after",
    )
    source_root = rooted(relocation["source_prefix_after"])
    source_files = archived_tree(source_root)
    assert source_files
    module = source_root / "src/xml_attributes.rs"
    census_tests = source_root / "src/census_tests_0798.rs"
    canonical_tests = source_root / "src/xml_attributes/tests.rs"
    assert module.is_file() and census_tests.is_file() and canonical_tests.is_file()
    module_text = module.read_text()
    census_text = census_tests.read_text()
    if number == 0:
        assert "census_drop_0798(&self.tag" in module_text
        assert "fn tag(content: &str)" in census_text
    else:
        assert "census_drop_0798(&self.tag" in module_text
        assert "fn make_tag(content: &str)" in census_text
    result.update(
        {
            "id": f"quality-failed-{number}",
            "stage": "isolated hook quality",
            "reason": relocation["reason"],
            "relocation": artifact(root / "relocation.json"),
            "source_archive": source_files,
        }
    )
    return result


def audit_build_failure() -> dict:
    root = p / "build-before-failed-0"
    relocation = c.read(root / "relocation.json")
    assert relocation["original"] == "build-before"
    assert relocation["retained"] == "build-before-failed-0"
    result = receipt_audit(
        root,
        relocation,
        [0, 0, 0, 101],
        {3: ["function `enable` is never used", "doc_lazy_continuation"]},
        "original",
        "retained",
        "commands.json",
    )
    probe_root = p / "probe-src-failed-0"
    probe_files = []
    for name, expected_hash in sorted(relocation["files"].items()):
        path = probe_root / name
        current = artifact(path)
        assert current["sha256"] == expected_hash
        probe_files.append(current)
    inputs = c.read(root / "inputs.json")
    for name, expected_hash in inputs["probe"].items():
        assert expected_hash == c.sha(probe_root / name)
    result.update(
        {
            "id": "build-before-failed-0",
            "stage": "control probe build",
            "reason": relocation["reason"],
            "relocation": artifact(root / "relocation.json"),
            "probe_archive": probe_files,
            "inputs": artifact(root / "inputs.json"),
            "source": artifact(root / "source.json"),
        }
    )
    return result


def audit_recovered_quality_and_build() -> dict:
    quality = c.read(p / "quality/complete.json")
    assert all(row["exit_code"] == 0 for row in quality["rows"])
    build = c.read(p / "build-before/build.json")
    assert all(row["exit_code"] == 0 for row in build["rows"])
    return {
        "quality": {
            "complete": artifact(p / "quality/complete.json"),
            "receipts": artifact(p / "quality/receipts.json"),
            "exit_codes": [row["exit_code"] for row in quality["rows"]],
        },
        "build_before": {
            "build": artifact(p / "build-before/build.json"),
            "commands": artifact(p / "build-before/commands.json"),
            "exit_codes": [row["exit_code"] for row in build["rows"]],
        },
    }


def result() -> dict:
    failures = [
        audit_lock_failure(),
        audit_quality_failure(0),
        audit_quality_failure(1),
        audit_build_failure(),
    ]
    return {
        "schema": "litchi.performance.0798.failure-audit.v1",
        "production_changed": False,
        "failed_attempts": failures,
        "recovery": audit_recovered_quality_and_build(),
        "path_policy": "receipt log paths are resolved through each retained relocation archive",
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
    print("0798 failure audit PASS")
