"""Finalize all 0806 target cleanup with one bounded deletion.

This is the root-owned cleanup entry point.  It reads only terminal build,
quality, capture, profile, and cross receipts; verifies every copied binary
identity while the target is live; records the four independent cleanup
witnesses; and removes the top-level target exactly once.  Packet reports,
logs, source archives, and audit inputs are never removed.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import custody as c


P = c.P
TARGET = c.TARGET
SUPPLEMENT = P / "amendment-preflight"
SUPPLEMENT_TARGET = TARGET / "amendment-preflight"
MAIN_KINDS = ("native", "allocation", "profile")
SETUP_KINDS = ("native", "allocation", "profile")
RECEIPT_KEYS = {"path", "bytes", "sha256"}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def packet_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label} is missing")
    return c.read(path)


def exact_descriptor(value: Any, expected: Path, parent: Path, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{label} descriptor is malformed")
    require(value.get("path") == str(expected), f"{label} path changed")
    require(expected.parent == parent, f"{label} escaped its owned target")
    require(expected.is_file() and not expected.is_symlink(),
            f"{label} binary is missing or symlinked")
    require(c.artifact(expected) == value, f"{label} identity changed")
    return dict(value)


def descriptor_file(value: Any, expected: Path, base: Path, label: str) -> Path:
    """Verify a packet descriptor whose path may be absolute or packet-relative."""

    require(isinstance(value, dict) and set(value) == RECEIPT_KEYS,
            f"{label} descriptor is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label} path is missing")
    path = Path(raw)
    path = path if path.is_absolute() else base / path
    require(path == expected, f"{label} path changed")
    require(path.is_file() and not path.is_symlink(), f"{label} is missing")
    actual = c.artifact(path)
    require(actual["bytes"] == value.get("bytes")
            and actual["sha256"] == value.get("sha256"),
            f"{label} identity changed")
    return path


def terminal_receipts(complete_path: Path, expected_children: int,
                      label: str, binary_ids: set[tuple[str, int, str]] | None = None,
                      expected_decodes: int | None = None,
                      base: Path = P) -> list[dict[str, Any]]:
    complete = packet_json(complete_path, f"{label} complete witness")
    require(isinstance(complete, dict), f"{label} complete witness is malformed")
    require(complete.get("children") == expected_children,
            f"{label} child count is not terminal")
    receipts = complete.get("receipts")
    receipts_path = descriptor_file(receipts, complete_path.parent / "receipts.json",
                                     base, f"{label}.receipts")
    rows = packet_json(receipts_path, f"{label} receipts")
    require(isinstance(rows, list) and len(rows) == expected_children,
            f"{label} receipt count is not terminal")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{label} receipt {index} did not succeed")
        if binary_ids is not None:
            binary = row.get("binary")
            require(isinstance(binary, dict), f"{label} receipt {index} has no binary")
            identity = (binary.get("path"), binary.get("bytes"), binary.get("sha256"))
            require(identity in binary_ids,
                    f"{label} receipt {index} references an unexpected binary")
    if expected_decodes is not None:
        require(complete.get("decodes") == expected_decodes,
                f"{label} decode count is not terminal")
        decode_desc = complete.get("decode_receipts")
        decode_path = descriptor_file(decode_desc, complete_path.parent / "decodes.json",
                                      base, f"{label}.decode_receipts")
        decoded = packet_json(decode_path, f"{label} decode receipts")
        require(isinstance(decoded, list) and len(decoded) == expected_decodes,
                f"{label} decode receipt count is not terminal")
        require(all(isinstance(row, dict) and row.get("exit_code") == 0
                    for row in decoded),
                f"{label} decode receipt failed")
    source = complete.get("source")
    if source is not None:
        descriptor_file(source, complete_path.parent / "source.json", base,
                        f"{label}.source")
    return rows


def check_quality() -> None:
    quality = packet_json(P / "quality.json", "final production quality")
    require(isinstance(quality, dict) and isinstance(quality.get("rows"), list)
            and len(quality["rows"]) == 6,
            "final production quality gate count changed")
    require(all(isinstance(row, dict) and row.get("exit_code") == 0
                for row in quality["rows"]),
            "final production quality is not terminal")

    checks = packet_json(P / "quality-2" / "checks.json", "final quality checks")
    require(isinstance(checks, list) and len(checks) == 6,
            "final quality check count changed")
    require(all(isinstance(row, dict) and row.get("exit_code") == 0
                for row in checks), "final quality check is not terminal")


def check_probe_quality(leg: str) -> None:
    root = P / f"probe-quality-{leg}"
    complete = packet_json(root / "complete.json", f"probe quality {leg}")
    require(isinstance(complete, dict), f"probe quality {leg} complete is malformed")
    descriptor_file(complete.get("receipts"), root / "receipts.json", P,
                    f"probe quality {leg}.receipts")
    rows = packet_json(root / "receipts.json", f"probe quality {leg} receipts")
    require(isinstance(rows, list) and len(rows) == 3,
            f"probe quality {leg} receipt count changed")
    require(all(isinstance(row, dict) and row.get("exit_code") == 0
                for row in rows), f"probe quality {leg} is not terminal")
    if leg == "after":
        descriptor_file(complete.get("inputs"), root / "inputs.json", P,
                        "probe quality after.inputs")


def collect_main() -> tuple[list[dict[str, Any]], dict[str, dict[str, dict[str, Any]]]]:
    removed: list[dict[str, Any]] = []
    builds: dict[str, dict[str, dict[str, Any]]] = {}
    for leg in ("before", "after"):
        build = packet_json(P / f"build-{leg}" / "build.json", f"build-{leg}")
        rows = build.get("rows")
        require(isinstance(rows, list) and len(rows) == 3,
                f"build-{leg} command count changed")
        require(all(isinstance(row, dict) and row.get("exit_code") == 0
                    for row in rows), f"build-{leg} is not terminal")
        binaries = build.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == set(MAIN_KINDS),
                f"build-{leg} binary set changed")
        builds[leg] = {}
        for kind in MAIN_KINDS:
            binary = exact_descriptor(binaries[kind], TARGET / f"{leg}-{kind}",
                                      TARGET, f"build-{leg}.{kind}")
            builds[leg][kind] = binary
            removed.append(binary)
    return removed, builds


def collect_cross() -> list[dict[str, Any]]:
    removed = []
    for leg in ("before", "after"):
        receipt = packet_json(P / f"cross-build-{leg}" / "receipt.json",
                              f"cross build {leg}")
        require(receipt.get("exit_code") == 0,
                f"cross build {leg} is not terminal")
        binary = exact_descriptor(receipt.get("binary"), TARGET / f"cross-{leg}",
                                  TARGET, f"cross build {leg}.binary")
        removed.append(binary)
    return removed


def collect_setup() -> list[dict[str, Any]]:
    removed = []
    for number in range(3):
        attempt = packet_json(P / f"setup-attempt-{number}.json",
                              f"setup attempt {number}")
        require(attempt.get("schema") == "litchi.performance.0806.setup-attempt.v1",
                f"setup attempt {number} schema changed")
        entries = attempt.get("retained_binaries")
        require(isinstance(entries, list) and len(entries) == 3,
                f"setup attempt {number} binary count changed")
        by_kind = {}
        for entry in entries:
            require(isinstance(entry, dict) and entry.get("kind") in SETUP_KINDS,
                    f"setup attempt {number} binary entry malformed")
            kind = entry["kind"]
            require(kind not in by_kind, f"setup attempt {number} binary duplicated")
            original = entry.get("original")
            retained = entry.get("retained")
            require(isinstance(original, dict) and isinstance(retained, dict),
                    f"setup attempt {number}.{kind} identity is incomplete")
            require(original.get("path") == str(TARGET / f"before-{kind}"),
                    f"setup attempt {number}.{kind} original path changed")
            require(retained.get("bytes") == original.get("bytes")
                    and retained.get("sha256") == original.get("sha256"),
                    f"setup attempt {number}.{kind} identity diverged")
            expected = TARGET / f"before-{kind}-setup-{number}"
            saved = exact_descriptor(retained, expected, TARGET,
                                     f"setup attempt {number}.{kind}")
            by_kind[kind] = saved
        require(set(by_kind) == set(SETUP_KINDS),
                f"setup attempt {number} binary kinds changed")
        removed.extend(by_kind[kind] for kind in SETUP_KINDS)
    return removed


def collect_supplement() -> tuple[list[dict[str, Any]], dict[str, Any]]:
    removed = []
    builds = {}
    for leg in ("before", "after"):
        build = packet_json(SUPPLEMENT / f"build-{leg}" / "build.json",
                            f"supplement build {leg}")
        require(build.get("schema") == "litchi.performance.0806.amendment-build.v1",
                f"supplement build {leg} schema changed")
        command = build.get("command")
        require(isinstance(command, dict) and command.get("exit_code") == 0,
                f"supplement build {leg} is not terminal")
        binary = exact_descriptor(build.get("binary"),
                                  SUPPLEMENT_TARGET / f"{leg}-native",
                                  SUPPLEMENT_TARGET,
                                  f"supplement build {leg}.binary")
        builds[leg] = build
        removed.append(binary)

    complete_path = SUPPLEMENT / "native" / "complete.json"
    complete = packet_json(complete_path, "supplement native")
    require(complete.get("schema") == "litchi.performance.0806.amendment-native.v1",
            "supplement native schema changed")
    rows = terminal_receipts(complete_path, 936, "supplement native",
                             base=SUPPLEMENT)
    require(complete.get("expected_children") == 936
            and complete.get("samples") == 28080
            and complete.get("profiles") is False
            and complete.get("callgrind") is False,
            "supplement native cardinality or profile policy changed")
    for index, row in enumerate(rows):
        require(row.get("leg") in ("before", "after"),
                f"supplement native receipt {index} leg is malformed")
    return removed, complete


def tree_bytes(root: Path) -> int:
    return sum(path.stat().st_size for path in root.rglob("*")
               if path.is_file() and not path.is_symlink())


def main() -> None:
    witnesses = [P / "cleanup.json", P / "cross-cleanup.json",
                 P / "setup-cleanup.json", SUPPLEMENT / "cleanup.json"]
    require(all(not path.exists() for path in witnesses),
            "cleanup witness already exists; refusing a second cleanup")
    require(TARGET.is_dir() and not TARGET.is_symlink(),
            f"owned target is missing or symlinked: {TARGET}")
    require(SUPPLEMENT_TARGET.is_dir() and not SUPPLEMENT_TARGET.is_symlink(),
            f"supplement target is missing or symlinked: {SUPPLEMENT_TARGET}")

    check_quality()
    check_probe_quality("before")
    check_probe_quality("after")
    main_removed, main_builds = collect_main()
    terminal_receipts(P / "qualification" / "complete.json", 18,
                      "qualification", {(main_builds["before"]["allocation"]["path"],
                                          main_builds["before"]["allocation"]["bytes"],
                                          main_builds["before"]["allocation"]["sha256"])})
    terminal_receipts(P / "native" / "complete.json", 216, "native",
                      {(main_builds[leg]["native"]["path"],
                        main_builds[leg]["native"]["bytes"],
                        main_builds[leg]["native"]["sha256"])
                       for leg in ("before", "after")})
    terminal_receipts(P / "allocation" / "complete.json", 72, "allocation",
                      {(main_builds[leg]["allocation"]["path"],
                        main_builds[leg]["allocation"]["bytes"],
                        main_builds[leg]["allocation"]["sha256"])
                       for leg in ("before", "after")})
    terminal_receipts(P / "profile" / "complete.json", 4, "profile",
                      {(main_builds[leg]["profile"]["path"],
                        main_builds[leg]["profile"]["bytes"],
                        main_builds[leg]["profile"]["sha256"])
                       for leg in ("before", "after")}, expected_decodes=8)
    cross_removed = collect_cross()
    cross_ids = {(item["path"], item["bytes"], item["sha256"])
                 for item in cross_removed}
    cross_before_id = (cross_removed[0]["path"], cross_removed[0]["bytes"],
                       cross_removed[0]["sha256"])
    terminal_receipts(P / "cross-qualification" / "complete.json", 1,
                      "cross qualification", {cross_before_id})
    terminal_receipts(P / "cross-native" / "complete.json", 12, "cross native",
                      cross_ids)
    supplement_removed, _supplement_complete = collect_supplement()

    all_sets = [main_removed, cross_removed, collect_setup(), supplement_removed]
    all_ids = [(item["path"], item["bytes"], item["sha256"])
               for group in all_sets for item in group]
    require(len(all_ids) == 19 and len(set(all_ids)) == 19,
            "cleanup descriptor sets overlap or changed cardinality")
    setup_removed = all_sets[2]
    removed_target_bytes = tree_bytes(TARGET)
    supplement_target_bytes = tree_bytes(SUPPLEMENT_TARGET)
    native_complete = {
        "path": "native/complete.json",
        "bytes": (SUPPLEMENT / "native" / "complete.json").stat().st_size,
        "sha256": c.sha(SUPPLEMENT / "native" / "complete.json"),
    }

    # This is deliberately the only deletion in this driver.  All packet
    # evidence stays under P, including the supplemental native receipts.
    shutil.rmtree(TARGET)
    require(not TARGET.exists(), "owned target was not removed")

    c.write(P / "cleanup.json", {
        "removed_binaries": main_removed,
        "removed_target_bytes": removed_target_bytes,
        "target": str(TARGET),
        "target_removed": True,
    })
    c.write(P / "cross-cleanup.json", {
        "schema": "litchi.performance.0806.cross-cleanup.v1",
        "target": str(TARGET),
        "binaries_removed": True,
        "removed_binaries": cross_removed,
    })
    c.write(P / "setup-cleanup.json", {
        "schema": "litchi.performance.0806.setup-cleanup.v1",
        "target": str(TARGET),
        "target_removed": True,
        "removed_binaries": setup_removed,
    })
    c.write(SUPPLEMENT / "cleanup.json", {
        "schema": "litchi.performance.0806.amendment-cleanup.v1",
        "target": str(SUPPLEMENT_TARGET),
        "target_removed": True,
        "removed_target_bytes": supplement_target_bytes,
        "removed_binaries": supplement_removed,
        "removed_failed_binaries": [],
        "native_complete": native_complete,
    })
    print("0806 cleanup complete: 6 main, 2 cross, 9 setup, 2 supplemental binaries")


if __name__ == "__main__":
    main()
