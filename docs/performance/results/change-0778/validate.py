"""Fail-closed offline validation for the 0778 durability packet."""

from __future__ import annotations

import ast
import json
import re
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET


def read_json(path: Path) -> Any:
    return analyze.read_json(path)


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def quality_commands(path: Path | None = None, expected_count: int = 9) -> list[list[str]]:
    """Evaluate only the literal command matrix in quality.py.

    The runner is intentionally not executed during replay.  Resolving its
    small AST keeps the receipt bound to the exact nine commands while still
    allowing the runner to construct shared flags in variables.
    """

    path = path or (PACKET / "quality.py")
    tree = ast.parse(path.read_text(), filename=str(path))
    assignments: dict[str, ast.AST] = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    assignments[target.id] = node.value
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            assignments[node.target.id] = node.value
    require("commands" in assignments, "quality.py command matrix is missing")

    values: dict[str, list[str]] = {}
    resolving: set[str] = set()

    def expand(node: ast.AST) -> list[str]:
        if isinstance(node, ast.Constant):
            require(isinstance(node.value, str), "quality command has a non-string literal")
            return [node.value]
        if isinstance(node, ast.Name):
            if node.id not in values:
                require(node.id in assignments, f"quality command references unknown {node.id}")
                require(node.id not in resolving,
                        f"quality command variable {node.id} is recursive")
                resolving.add(node.id)
                values[node.id] = expand(assignments[node.id])
                resolving.remove(node.id)
            return list(values[node.id])
        if isinstance(node, (ast.List, ast.Tuple)):
            result: list[str] = []
            for item in node.elts:
                if isinstance(item, ast.Starred):
                    result.extend(expand(item.value))
                else:
                    result.extend(expand(item))
            return result
        if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
            return expand(node.left) + expand(node.right)
        # Some quality runners build an environment or target path but the
        # command list itself must remain literal.  Rejecting this is safer
        # than accepting an unbound command.
        fail = getattr(analyze, "fail")
        fail(f"quality command matrix contains unsupported AST: {ast.dump(node)}")
        return []

    node = assignments["commands"]
    require(isinstance(node, ast.List), "quality.py commands is not a list")
    matrix = [expand(item) for item in node.elts]
    require(len(matrix) == expected_count, "quality command count changed")
    return matrix


def check_quality_commands() -> None:
    _, _, rows, _ = analyze.quality_receipt()
    require(len(rows) == 9, "quality rows are incomplete")
    expected = quality_commands()
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"quality row {index} is invalid")
        command = row.get("command")
        require(command == expected[index], f"quality command {index + 1} differs from quality.py")


def check_docx_quality_commands() -> None:
    quality = analyze.check_docx_quality()
    expected = quality_commands(PACKET / "docx_quality.py", 5)
    for index, row in enumerate(quality["commands"]):
        require(row["command"] == expected[index],
                f"DOCX owner quality command {index + 1} differs from docx_quality.py")


def check_source_custody(plan: dict[str, Any], build: dict[str, Any]) -> None:
    source = analyze.load_source_custody(plan)
    require(source is not None, "source custody manifest is missing")
    build_source = build.get("source")
    if isinstance(build_source, dict) and isinstance(build_source.get("files"), dict):
        build_source = build_source["files"]
    if build_source is None and isinstance(build.get("source_files"), dict):
        build_source = build["source_files"]
    if isinstance(build_source, dict):
        require(source == build_source, "build/source manifest differs from retained source custody")
    if "source_sha256" in build:
        require(build["source_sha256"] == analyze.sha256(PACKET / "source.json"),
                "build source digest differs from packet source custody")
    for name, digest in source.items():
        require(analyze.is_sha(digest), f"source custody digest is invalid: {name}")

    quality_path, _, _, _ = analyze.quality_receipt()
    quality_source = quality_path.parent / "source.json"
    if quality_source.is_file():
        require(read_json(quality_source) == read_json(PACKET / "source.json"),
                "successful quality source census differs from packet custody")

    # If the capture has per-child source witnesses, require both sides and
    # exact byte identity.  This also catches an archived absolute path that
    # silently resolves to a different checkout.
    for directory_name in ("native-0", "allocation-0", "qualification-0", "capture-0"):
        directory = PACKET / directory_name
        if not directory.is_dir():
            continue
        for path in sorted(directory.glob("**/*source-before*.json")):
            after = path.with_name(path.name.replace("source-before", "source-after"))
            require(after.is_file(), f"missing source-after witness for {path.name}")
            before_value = read_json(path)
            after_value = read_json(after)
            require(before_value == after_value, f"source changed during child: {path.name}")


def check_binary_cleanup(build: dict[str, Any]) -> None:
    cleanup_path = PACKET / "cleanup.json"
    cleanup = read_json(cleanup_path) if cleanup_path.is_file() else None
    cleanup_verified = isinstance(cleanup, dict) and (
        cleanup.get("verified") is True
        or cleanup.get("executables_verified_before_removal") is True
    )
    binaries = build.get("binaries")
    require(isinstance(binaries, dict), "build binary map is missing")
    for name, receipt in binaries.items():
        path = analyze.resolve_path(receipt.get("path"), capture_bound=False)
        if path.is_file():
            analyze.artifact_receipt(receipt, f"binary {name}", capture_bound=False)
            continue
        require(cleanup_verified, f"binary {name} is missing without cleanup verification")
        # A cleanup record must carry the exact path, byte count and digest;
        # a bare 'removed: true' flag is not custody evidence.
        found = False

        def walk(value: Any) -> None:
            nonlocal found
            if found:
                return
            if isinstance(value, dict):
                if value.get("path") == receipt.get("path") and value.get("sha256") == receipt.get("sha256"):
                    if value.get("bytes", value.get("size")) == receipt.get("bytes", receipt.get("size")):
                        found = True
                for child in value.values():
                    walk(child)
            elif isinstance(value, list):
                for child in value:
                    walk(child)

        walk(cleanup)
        require(found, f"binary {name} lacks exact cleanup witness")


def check_capture_runner_hashes() -> None:
    runner_receipt = PACKET / "runner.json"
    if runner_receipt.is_file():
        value = read_json(runner_receipt)
        analyze.artifact_receipt(value, "capture runner receipt", capture_bound=True)
        require(value.get("sha256") == analyze.sha256(PACKET / "capture.py"),
                "capture runner receipt changed")
    freeze_path = PACKET / "freeze.json"
    if freeze_path.is_file():
        freeze = read_json(freeze_path)
        require(isinstance(freeze, dict), "freeze.json is malformed")
        for key, path in (("plan", PACKET / "plan.json"),
                          ("source", PACKET / "source.json"),
                          ("build", PACKET / "build.json"),
                          ("runner", PACKET / "capture.py")):
            receipt = freeze.get(key)
            if receipt is None:
                continue
            actual = analyze.artifact_receipt(receipt, f"freeze {key}", capture_bound=True)
            assert actual is not None
            require(analyze.sha256(actual) == analyze.sha256(path),
                    f"freeze {key} is not current-source bound")
    for directory_name in ("native-0", "allocation-0", "qualification-0", "capture-0"):
        directory = PACKET / directory_name
        complete_path = directory / "complete.json"
        if not complete_path.is_file():
            continue
        complete = read_json(complete_path)
        runner_sha = complete.get("runner_sha256", complete.get("script_sha256"))
        if runner_sha is None:
            continue
        runner = analyze.find_packet_file(("capture.py",))
        require(runner is not None and analyze.sha256(runner) == runner_sha,
                f"{directory_name} runner hash changed")
        plan_sha = complete.get("plan_sha256")
        if plan_sha is not None:
            require(plan_sha == analyze.sha256(PACKET / "plan.json"),
                    f"{directory_name} plan hash changed")


def check_optional_seal() -> None:
    seal_path = PACKET / "seal.json"
    if not seal_path.is_file():
        return
    seal = read_json(seal_path)
    require(isinstance(seal, dict) and isinstance(seal.get("files"), dict),
            "seal.json is malformed")
    actual = {
        str(path.relative_to(PACKET)): analyze.sha256(path)
        for path in sorted(PACKET.rglob("*"))
        if path.is_file() and path != seal_path
    }
    require(actual == seal["files"], "packet seal does not match current bytes")


def check_analysis_replay(result: dict[str, Any]) -> None:
    path = PACKET / "analysis.json"
    require(path.is_file(), "analysis.json is missing; run analyze.py after capture")
    require(read_json(path) == result, "analysis.json does not replay byte-for-byte")


def validate() -> dict[str, Any]:
    plan = analyze.load_plan()
    quality = analyze.check_quality()
    build, cleanup = analyze.load_build_and_cleanup()
    check_quality_commands()
    check_docx_quality_commands()
    check_source_custody(plan, build)
    check_binary_cleanup(build)
    check_capture_runner_hashes()
    result = analyze.analyze()
    check_analysis_replay(result)
    check_optional_seal()

    # The analysis must retain the exact process cardinalities specified by
    # the frozen plan.  These checks are intentionally independent of the
    # number of summary groups, so a dropped control cannot hide in a group.
    require(result["native"]["children"] == plan["expected_children"]["native"],
            "analysis native cardinality differs from plan")
    require(result["allocation"]["children"] == plan["expected_children"]["allocation"],
            "analysis allocation cardinality differs from plan")
    require(result["qualification"]["children"] == plan["expected_children"]["qualification"],
            "analysis qualification cardinality differs from plan")
    require(result["trace"]["children"] == plan["expected_children"]["qualification"],
            "analysis trace cardinality differs from plan")
    require(result["verification"]["cleanup_binary_witness_required_when_missing"] is True,
            "analysis cleanup custody contract is not stable")
    return {
        "quality_gates": quality["gates"],
        "quality_counts": quality["counts"],
        "native_children": result["native"]["children"],
        "allocation_children": result["allocation"]["children"],
        "qualification_children": result["qualification"]["children"],
        "trace_children": result["trace"]["children"],
        "spread_flags": len(result["native"]["analysis"]["spread_flags_over_5_percent"]),
        "allocation_spread_flags": len(result["allocation"]["analysis"]["spread_flags_over_5_percent"]),
        "cleanup_verified": cleanup,
        "optional_seal": (PACKET / "seal.json").is_file(),
    }


if __name__ == "__main__":
    print(json.dumps(validate(), indent=2, sort_keys=True))
