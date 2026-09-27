"""Fail-closed offline validation for the 0777 capture packet."""

from __future__ import annotations

import json
import re
import ast
from pathlib import Path
from typing import Any

import analyze

try:
    import observe
except ModuleNotFoundError:
    # Keep import-time syntax checks useful while the separately owned
    # observer is still being prepared.  Validation remains fail-closed until
    # both the module and its receipt exist.
    observe = None  # type: ignore[assignment]


PACKET = analyze.PACKET
CAPTURE = analyze.CAPTURE


def read_json(path: Path) -> Any:
    return analyze.read_json(path)


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def source_mapping(value: Any, label: str) -> dict[str, str]:
    if isinstance(value, dict) and isinstance(value.get("files"), dict):
        value = value["files"]
    require(isinstance(value, dict), f"{label} is not a source mapping")
    result: dict[str, str] = {}
    for name, digest in value.items():
        require(isinstance(name, str), f"{label} has a non-string path")
        require(analyze.is_hex_digest(digest), f"{label} has invalid SHA for {name}")
        result[name] = digest
    return result


def quality_command_matrix() -> list[list[str]]:
    """Read quality.py's literal command matrix without executing it.

    The quality runner keeps its matrix inside ``__main__``.  Evaluating the
    small list expressions from its AST binds this validator to the runner's
    actual package order and shared flags while avoiding a second quality
    workload during offline replay.
    """

    path = PACKET / "quality.py"
    source = path.read_text()
    try:
        tree = ast.parse(source, filename=str(path))
    except SyntaxError as error:
        raise RuntimeError(f"quality.py is not parseable: {error}") from error
    assignments: dict[str, ast.AST] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    assignments[target.id] = node.value
    require("packages" in assignments, "quality.py package matrix is missing")
    require("common" in assignments, "quality.py common flags are missing")
    require("commands" in assignments, "quality.py command matrix is missing")

    def constants(node: ast.AST, name: str) -> list[str]:
        require(isinstance(node, (ast.List, ast.Tuple)), f"quality.py {name} is not a list")
        values: list[str] = []
        for element in node.elts:
            require(isinstance(element, ast.Constant) and isinstance(element.value, str), f"quality.py {name} has a non-string")
            values.append(element.value)
        return values

    packages = constants(assignments["packages"], "packages")
    common = constants(assignments["common"], "common")
    owners = [argument for package in packages for argument in ("-p", package)]
    values: dict[str, list[str]] = {
        "packages": packages,
        "common": common,
        "owners": owners,
    }

    def expand(node: ast.AST) -> list[str]:
        if isinstance(node, ast.Constant):
            require(isinstance(node.value, str), "quality.py command has a non-string literal")
            return [node.value]
        if isinstance(node, ast.Name):
            require(node.id in values, f"quality.py command references unknown {node.id}")
            return list(values[node.id])
        if isinstance(node, (ast.List, ast.Tuple)):
            result: list[str] = []
            for element in node.elts:
                if isinstance(element, ast.Starred):
                    result.extend(expand(element.value))
                else:
                    result.extend(expand(element))
            return result
        raise RuntimeError(f"quality.py command matrix contains unsupported AST: {ast.dump(node)}")

    matrix_node = assignments["commands"]
    require(isinstance(matrix_node, ast.List), "quality.py commands is not a list")
    matrix = [expand(command) for command in matrix_node.elts]
    require(len(matrix) == 9, "quality.py command matrix must contain 9 commands")
    return matrix


def check_quality() -> tuple[dict[str, Any], dict[str, str], dict[str, int]]:
    path = PACKET / "quality.json"
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is not an object")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == 9, "quality receipt must contain 9 gates")
    expected_commands = quality_command_matrix()
    source_name = quality.get("source_file")
    require(isinstance(source_name, str), "quality receipt has no source file")
    source_path = analyze.resolve_packet_path(source_name)
    require(source_path.is_file(), f"missing quality source receipt: {source_name}")
    source_digest = quality.get("source_sha256")
    require(analyze.is_hex_digest(source_digest), "quality source SHA is invalid")
    require(analyze.sha256(source_path) == source_digest, "quality source receipt changed")
    source = source_mapping(read_json(source_path), "quality source")
    counts = {"passed": 0, "ignored": 0, "suites": 0}
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"quality row {index} is not an object")
        require(
            row.get("command") == expected_commands[index],
            f"quality command {index + 1} differs from quality.py",
        )
        require(row.get("exit") == 0, f"quality gate {index + 1} failed")
        log = row.get("log")
        if isinstance(log, dict):
            log_path = analyze.check_receipt(
                log, f"quality gate {index + 1} log", capture_bound=False
            )
            assert log_path is not None
        else:
            require(isinstance(log, str), f"quality gate {index + 1} has no log")
            log_path = analyze.resolve_packet_path(log)
            require(log_path.is_file(), f"missing quality log: {log}")
            require(analyze.is_hex_digest(row.get("sha256")), f"quality gate {index + 1} log SHA is invalid")
            require(analyze.sha256(log_path) == row["sha256"], f"quality gate {index + 1} log changed")
        text = log_path.read_text(errors="replace")
        for match in re.finditer(
            r"test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored",
            text,
        ):
            status, passed, failed, ignored = match.groups()
            require(status == "ok" and int(failed) == 0, f"quality test failure in gate {index + 1}")
            counts["passed"] += int(passed)
            counts["ignored"] += int(ignored)
            counts["suites"] += 1
    return quality, source, counts


def check_source_receipts(
    complete: dict[str, Any], quality_source: dict[str, str]
) -> dict[str, str]:
    final = read_json(PACKET / "final-source.json")
    require(isinstance(final, dict), "final-source.json is not an object")
    final_files = source_mapping(final, "final-source")
    require(len(final_files) == 305, f"final-source must bind 305 files, found {len(final_files)}")
    candidate = complete.get("candidate")
    require(final.get("candidate") == candidate, "final-source candidate differs from capture")
    origin = read_json(PACKET / "origin.json")
    main_base = origin.get("main_base")
    require(isinstance(main_base, str), "origin receipt has no main base")
    require(final.get("base") == main_base[:10], "final-source base differs from origin")
    for name, digest in final_files.items():
        require(quality_source.get(name) == digest, f"quality source disagrees for {name}")

    source_after = read_json(CAPTURE / "source-after.json")
    source_before = read_json(CAPTURE / "source-before.json")
    source_final = read_json(CAPTURE / "source-final.json")
    require(source_after.get("head") == candidate, "source-after head differs from candidate")
    require(
        isinstance(source_before.get("head"), str)
        and source_before["head"].startswith(main_base),
        "source-before head differs from origin",
    )
    after_files = source_mapping(source_after, "capture source-after")
    before_files = source_mapping(source_before, "capture source-before")
    for name, digest in final_files.items():
        require(after_files.get(name) == digest, f"source-after disagrees for {name}")
    require(source_final.get("after") == source_after, "source-final after receipt changed")
    require(source_final.get("before") == source_before, "source-final before receipt changed")
    # The candidate source map must cover every file bound by final-source.
    # The base predates the 0770 harness additions, so it may legitimately
    # omit a few of those 305 candidate/harness paths.
    for name in final_files:
        require(name in after_files, f"source-after omitted {name}")
    return final_files


def check_probe_and_lock_receipts() -> None:
    inputs = read_json(CAPTURE / "probe-inputs.json")
    require(isinstance(inputs, dict), "probe-inputs.json is not an object")
    require(set(inputs) == set(analyze.PROBE_INPUTS), "probe input set changed")
    for name in analyze.PROBE_INPUTS:
        receipt = inputs[name]
        require(isinstance(receipt, dict), f"probe input {name} is invalid")
        source_name = receipt.get("path")
        source_path = analyze.resolve_packet_path(source_name)
        require(source_path.is_file(), f"missing archived probe input {name}")
        expected_sha = receipt.get("sha256")
        require(analyze.is_hex_digest(expected_sha), f"probe input {name} SHA is invalid")
        require(analyze.sha256(source_path) == expected_sha, f"probe input {name} changed")
        require(receipt.get("origin_sha256") == expected_sha, f"probe input {name} origin changed")
        require(receipt.get("bytes") == source_path.stat().st_size, f"probe input {name} size changed")

    lock = read_json(CAPTURE / "standalone-lock.json")
    require(isinstance(lock, dict), "standalone-lock.json is not an object")
    before_lock = analyze.check_receipt(lock.get("before"), "standalone before lock")
    after_lock = analyze.check_receipt(lock.get("after"), "standalone after lock")
    require(before_lock is not None and after_lock is not None, "standalone lock is missing")
    require(analyze.sha256(before_lock) == analyze.sha256(after_lock), "standalone lock differs between legs")
    require(lock.get("copied_before_to_after") is True, "standalone lock was not shared")

    for leg in ("before", "after"):
        mapping = read_json(CAPTURE / f"source-map-{leg}.json")
        require(isinstance(mapping, dict), f"source-map-{leg} is not an object")
        files = mapping.get("files")
        require(isinstance(files, dict), f"source-map-{leg} has no files")
        require(
            mapping.get("template_sha256") == inputs["Cargo.toml.template"]["sha256"],
            f"source-map-{leg} template changed",
        )
        for destination, copied in files.items():
            require(isinstance(copied, dict), f"source-map-{leg} {destination} is invalid")
            target = CAPTURE / f"probe-{leg}" / destination
            require(target.is_file(), f"missing copied probe {leg}/{destination}")
            require(analyze.is_hex_digest(copied.get("sha256")), f"copied probe {destination} has bad SHA")
            require(analyze.sha256(target) == copied["sha256"], f"copied probe {leg}/{destination} changed")
            source_name = copied.get("source")
            if destination != "Cargo.toml":
                require(source_name in inputs, f"unknown copied probe source {source_name}")
                require(copied["sha256"] == inputs[source_name]["sha256"], f"copied probe source mismatch: {destination}")
        copied_lock = CAPTURE / f"probe-{leg}" / "Cargo.lock"
        require(copied_lock.is_file(), f"missing copied Cargo.lock for {leg}")
        require(analyze.sha256(copied_lock) == analyze.sha256(before_lock), f"probe lock differs for {leg}")


def check_inventory_guards(complete: dict[str, Any]) -> None:
    for field in (
        "source_unchanged",
        "probe_inputs_unchanged",
        "fixtures_unchanged",
        "package_inventory_unchanged",
        "standalone_lock_shared",
        "binaries_unchanged",
    ):
        require(complete.get(field) is True, f"capture guard {field} is not true")
    runner = read_json(CAPTURE / "runner.json")
    runner_path = analyze.check_receipt(runner, "capture runner", capture_bound=False)
    require(runner_path == PACKET / "capture.py", "capture runner receipt did not relocate to packet")
    for before_name, after_name in (
        ("fixtures-before.json", "fixtures-after.json"),
        ("package-inventory-before.json", "package-inventory-after.json"),
    ):
        before = read_json(CAPTURE / before_name)
        after = read_json(CAPTURE / after_name)
        require(before == after, f"capture inventory changed: {before_name}")


def check_cleanup_receipt() -> bool:
    """Validate the optional post-capture target cleanup receipt."""

    path = PACKET / "cleanup.json"
    if not path.is_file():
        return False
    cleanup = read_json(path)
    require(isinstance(cleanup, dict), "cleanup.json is not an object")
    for field in ("source_unchanged", "fixtures_unchanged"):
        require(cleanup.get(field) is True, f"cleanup guard {field} is not true")
    if "package_inventory_unchanged" in cleanup:
        require(cleanup["package_inventory_unchanged"] is True, "cleanup package inventory guard is not true")
    if "executables_verified_before_removal" in cleanup:
        require(cleanup["executables_verified_before_removal"] is True, "cleanup executable guard is not true")
    targets = cleanup.get("targets")
    require(isinstance(targets, list) and len(targets) == 3, "cleanup target receipt is incomplete")
    expected = {
        "/home/zhuhe/code/litchi-target-0777-quality",
        "/home/zhuhe/code/litchi-target-0777-release-before",
        "/home/zhuhe/code/litchi-target-0777-release-after",
    }
    seen: set[str] = set()
    for index, target in enumerate(targets):
        require(isinstance(target, dict), f"cleanup target {index} is invalid")
        name = target.get("path")
        require(name in expected, f"unexpected cleanup target: {name!r}")
        require(name not in seen, f"duplicate cleanup target: {name}")
        require(target.get("removed") is True, f"cleanup target was not removed: {name}")
        require(isinstance(target.get("bytes"), int) and target["bytes"] >= 0, f"cleanup target size is invalid: {name}")
        seen.add(name)
    require(seen == expected, "cleanup target set is incomplete")
    return True


def check_binary_receipts(runs: list[dict[str, Any]], cleanup: bool) -> None:
    binaries = read_json(CAPTURE / "binaries.json")
    require(isinstance(binaries, dict), "binaries.json is not an object")
    expected = {
        "before": ("attribute_checks", "mce-stream-probe"),
        "after": ("attribute_checks", "attribute_checks_equivalence", "mce-stream-probe"),
    }
    missing = False
    for leg, names in expected.items():
        require(set(binaries.get(leg, {})) == set(names), f"binary set changed for {leg}")
        for name in names:
            path = analyze.check_receipt(
                binaries[leg][name],
                f"{leg}/{name} binary",
                capture_bound=False,
                allow_missing=True,
            )
            missing = missing or path is None
    for index, row in enumerate(runs):
        binary = row.get("binary")
        if not isinstance(binary, dict):
            continue
        leg = row.get("leg")
        name = Path(str(binary.get("path", ""))).name
        if leg in binaries and name in binaries[leg]:
            require(binary == binaries[leg][name], f"run {index} binary receipt differs")
    require(not missing or cleanup, "binary target is absent without a validated cleanup receipt")


def check_extra_capture_receipts() -> None:
    native_receipts = CAPTURE / "native-receipts.json"
    if native_receipts.is_file():
        rows = read_json(native_receipts)
        require(isinstance(rows, list) and len(rows) == analyze.EXPECTED_NATIVE_PROCESSES, "native receipt set incomplete")
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"native receipt {index} is invalid")
            analyze.check_receipt(row.get("report"), f"native receipt {index} report")
            analyze.check_receipt(row.get("rss"), f"native receipt {index} RSS")
    for path in CAPTURE.glob("*-receipt.json"):
        value = read_json(path)
        if isinstance(value, dict) and "path" in value:
            analyze.check_receipt(value, path.name)
        elif isinstance(value, dict) and isinstance(value.get("report"), dict):
            analyze.check_receipt(value["report"], f"{path.name} report")


def check_seal() -> None:
    seal_path = PACKET / "seal.json"
    if not seal_path.is_file():
        return
    seal = read_json(seal_path)
    require(isinstance(seal, dict) and isinstance(seal.get("files"), dict), "invalid packet seal")
    actual = {
        str(path.relative_to(PACKET)): analyze.sha256(path)
        for path in sorted(PACKET.rglob("*"))
        if path.is_file() and path != seal_path
    }
    require(actual == seal["files"], "packet seal does not match current bytes")


def check_observations() -> dict[str, Any]:
    require(observe is not None, "observer module is missing")
    path = PACKET / "observations.json"
    recorded = read_json(path)
    observed = observe.observe()  # type: ignore[union-attr]
    require(recorded == observed, "observations.json does not replay exactly")
    require(isinstance(recorded, dict), "observations.json is not an object")
    return recorded


def check_operation_bounds(final_files: dict[str, str]) -> None:
    bounds = read_json(PACKET / "operation-bounds.json")
    log = PACKET / bounds["test_log"]
    require(analyze.sha256(log) == bounds["test_log_sha256"], "comparison-bound log changed")
    require(final_files.get(bounds["source"]) == bounds["source_sha256"], "comparison-bound source changed")
    text = log.read_text()
    expected = {
        "a_tag_of_distinct_names_costs_n_log_n_comparisons_in_any_order",
        "a_late_duplicate_costs_n_log_n_comparisons",
        "comparisons_grow_by_about_the_factor_of_names",
        "quick_xml_checks_the_first_names_and_this_iterator_the_rest",
    }
    require(set(bounds["passing_copies_per_test"]) == expected, "comparison-bound test set changed")
    for name, count in bounds["passing_copies_per_test"].items():
        matches = re.findall(r"test [^\n]*::" + re.escape(name) + r" \.\.\. ok", text)
        require(count == 5 and len(matches) == count, f"comparison-bound test missing: {name}")


def validate() -> dict[str, Any]:
    complete, runs = analyze.load_complete()
    quality, quality_source, counts = check_quality()
    final_files = check_source_receipts(complete, quality_source)
    check_operation_bounds(final_files)
    check_probe_and_lock_receipts()
    check_inventory_guards(complete)
    cleanup = check_cleanup_receipt()
    check_binary_receipts(runs, cleanup)
    check_extra_capture_receipts()
    observations = check_observations()
    result = analyze.analyze()
    analysis_path = PACKET / "analysis.json"
    require(analysis_path.is_file(), "analysis.json is missing; run analyze.py after capture completion")
    require(read_json(analysis_path) == result, "analysis.json does not replay exactly")
    require(len(result["cases"]) == len(analyze.EXPECTED_CASES), "analysis case count is incomplete")
    require(result["differential"]["changes"] == {}, "analysis contains differential changes")
    require(result["equivalence"]["mismatches"] == 0, "analysis contains equivalence mismatches")
    check_seal()
    return {
        "quality_gates": len(quality["rows"]),
        "quality_tests": counts,
        "source_files": len(final_files),
        "native_processes": analyze.EXPECTED_NATIVE_PROCESSES,
        "cases": len(result["cases"]),
        "differential_changes": len(result["differential"]["changes"]),
        "equivalence_mismatches": result["equivalence"]["mismatches"],
        "observations": {
            "instruction_processes": sum(
                item["rows"] + item["iterator_rows"]
                for item in observations["instructions"].values()
            ),
            "allocation_processes": sum(
                item["rows"] for item in observations["allocations"].values()
            ),
        },
        "binary_live_files_not_required": cleanup,
    }


if __name__ == "__main__":
    print(json.dumps(validate(), indent=2, sort_keys=True))
