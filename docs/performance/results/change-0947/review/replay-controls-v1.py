#!/usr/bin/env python3
"""Moved-input replay and rejection controls for frozen span-review v4."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any

TOOL_SHA = "d8dbc469e504f16aab022d2e3115dda2b3746f10a1300d0f2996ddb6cea93bb5"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def copy_inputs(repo: Path, identity_src: Path, dest: Path) -> None:
    if dest.exists():
        shutil.rmtree(dest)
    dest.mkdir(parents=True)
    identity = read_json(identity_src)
    identity_dst = dest / "identity/verified-v1.json"
    identity_dst.parent.mkdir(parents=True)
    shutil.copy2(identity_src, identity_dst)
    manifest_src = repo / "docs/performance/results/change-0904/artifacts/manifest.json"
    manifest_dst = dest / "docs/performance/results/change-0904/artifacts/manifest.json"
    manifest_dst.parent.mkdir(parents=True)
    shutil.copy2(manifest_src, manifest_dst)
    for case in identity["cases"]:
        for part in case["parts"]:
            for side in ("original", "replacement"):
                relative = Path(part[side]["path"])
                source = repo / "docs/performance/results/change-0904/artifacts" / relative
                target = dest / "docs/performance/results/change-0904/artifacts" / relative
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)


def tree_manifest(root: Path) -> list[dict[str, Any]]:
    entries = []
    for path in sorted(root.rglob("*")):
        if path.is_file() and not path.is_symlink():
            relative = path.relative_to(root).as_posix()
            entries.append({"path": relative, "bytes": path.stat().st_size,
                            "sha256": sha256(path)})
    return entries


def run_tool(tool: Path, repo: Path, identity: Path, output: Path) -> dict[str, Any]:
    argv = [sys.executable, "-B", str(tool), "--repo", str(repo),
            "--identity", str(identity), "--output", str(output)]
    env = dict(os.environ)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    result = subprocess.run(argv, cwd=repo, env=env, text=True,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            check=False)
    return {"argv": argv, "cwd": str(repo), "returncode": result.returncode,
            "stdout": result.stdout, "stderr": result.stderr,
            "output": str(output), "output_exists": output.exists()}


def normalize_projection(path: Path) -> dict[str, Any]:
    value = read_json(path)
    value["identity_receipt"]["path"] = "<normalized-identity-path>"
    return value


def mutate_identity_hash(identity_path: Path) -> None:
    value = read_json(identity_path)
    value["cases"][0]["parts"][0]["original"]["sha256"] = "0" * 64
    identity_path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def mutate_duplicate_cell(repo: Path, identity_path: Path) -> None:
    identity = read_json(identity_path)
    part = identity["cases"][0]["parts"][1]  # medium one-percent sheet1 original
    path = repo / "docs/performance/results/change-0904/artifacts" / part["original"]["path"]
    data = path.read_bytes()
    marker = b"</row>"
    duplicate = b'<c r="A1"><v>0</v></c>'
    if marker not in data:
        raise AssertionError("worksheet has no row close marker")
    mutated = data.replace(marker, duplicate + marker, 1)
    path.write_bytes(mutated)
    part["original"]["bytes"] = len(mutated)
    part["original"]["sha256"] = sha256(path)
    identity_path.write_text(json.dumps(identity, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--identity", type=Path, required=True)
    parser.add_argument("--tool", type=Path, required=True)
    parser.add_argument("--review-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.resolve()
    identity = args.identity.resolve()
    tool = args.tool.resolve()
    review = args.review_root.resolve()
    work = review / "replay-controls-work-v1"
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True)
    if sha256(tool) != TOOL_SHA:
        raise AssertionError("frozen v4 tool hash differs")

    original = work / "original-inputs"
    moved = work / "moved-inputs"
    copy_inputs(repo, identity, original)
    copy_inputs(repo, identity, moved)
    original_identity = original / "identity/verified-v1.json"
    moved_identity = moved / "identity/verified-v1.json"
    original_output = work / "original-replay.json"
    moved_output = work / "moved-replay.json"
    original_run = run_tool(tool, original, original_identity, original_output)
    moved_run = run_tool(tool, moved, moved_identity, moved_output)
    if original_run["returncode"] != 0 or moved_run["returncode"] != 0:
        raise AssertionError("positive replay failed")
    original_projection = normalize_projection(original_output)
    moved_projection = normalize_projection(moved_output)
    if original_projection != moved_projection:
        raise AssertionError("moved semantic projection differs")
    original_tree = tree_manifest(original)
    moved_tree = tree_manifest(moved)
    if original_tree != moved_tree:
        raise AssertionError("moved input tree manifest differs")

    negatives: dict[str, Any] = {}
    case = work / "corrupt-input"
    copy_inputs(repo, identity, case)
    corrupt = case / "docs/performance/results/change-0904/artifacts/medium/one-percent/original/xl/worksheets/sheet1.xml"
    data = corrupt.read_bytes()
    corrupt.write_bytes(data[:100] + bytes([data[100] ^ 1]) + data[101:])
    run = run_tool(tool, case, case / "identity/verified-v1.json", case / "output.json")
    if run["returncode"] == 0 or run["output_exists"]:
        raise AssertionError("corrupt input was accepted")
    negatives["corrupt_input_descriptor"] = run

    case = work / "missing-member"
    copy_inputs(repo, identity, case)
    missing = case / "docs/performance/results/change-0904/artifacts/medium/one-percent/replacement/xl/worksheets/sheet4.xml"
    missing.unlink()
    run = run_tool(tool, case, case / "identity/verified-v1.json", case / "output.json")
    if run["returncode"] == 0 or run["output_exists"]:
        raise AssertionError("missing member was accepted")
    negatives["missing_member"] = run

    case = work / "wrong-identity-binding"
    copy_inputs(repo, identity, case)
    mutate_identity_hash(case / "identity/verified-v1.json")
    run = run_tool(tool, case, case / "identity/verified-v1.json", case / "output.json")
    if run["returncode"] == 0 or run["output_exists"]:
        raise AssertionError("wrong identity hash was accepted")
    negatives["wrong_identity_hash"] = run

    case = work / "duplicate-cell-key"
    copy_inputs(repo, identity, case)
    duplicate_identity = case / "identity/verified-v1.json"
    mutate_duplicate_cell(case, duplicate_identity)
    run = run_tool(tool, case, duplicate_identity, case / "output.json")
    if run["returncode"] == 0 or run["output_exists"]:
        raise AssertionError("duplicate cell key was accepted")
    negatives["duplicate_cell_key"] = run

    case = work / "overwrite"
    copy_inputs(repo, identity, case)
    overwrite_output = case / "output.json"
    first = run_tool(tool, case, case / "identity/verified-v1.json", overwrite_output)
    second = run_tool(tool, case, case / "identity/verified-v1.json", overwrite_output)
    if first["returncode"] != 0 or second["returncode"] == 0:
        raise AssertionError("output overwrite control failed")
    negatives["output_overwrite"] = {"first": first, "second": second}

    receipt = {
        "schema": "litchi.xlsx-changed-cell-span-replay-controls-0947.v1",
        "status": "pass",
        "scope": "Moved copied-input replay and offline rejection controls for frozen v4; no workload/build/Rust/native measurement.",
        "command": "PYTHONDONTWRITEBYTECODE=1 python3 -B replay-controls-v1.py --repo /home/zhuhe/code/litchi --identity /home/zhuhe/litchi-performance-0947/identity/verified-v1.json --tool /home/zhuhe/litchi-performance-0947/review/changed-cell-span-review-v4.py --review-root /home/zhuhe/litchi-performance-0947/review --output /home/zhuhe/litchi-performance-0947/review/replay-controls-v1.json",
        "tool": {"path": str(tool), "bytes": tool.stat().st_size, "sha256": sha256(tool)},
        "identity": {"path": str(identity), "bytes": identity.stat().st_size, "sha256": sha256(identity)},
        "positive_replay": {
            "original": original_run, "moved": moved_run,
            "input_tree_byte_manifest_equal": True,
            "original_output_bytes": original_output.stat().st_size,
            "moved_output_bytes": moved_output.stat().st_size,
            "raw_output_byte_equal": original_output.read_bytes() == moved_output.read_bytes(),
            "semantic_projection_equal": True,
            "identity_path_normalization": "only identity_receipt.path is normalized; raw output differs solely because that field is absolute",
        },
        "negative_controls": {
            name: {"rejected": (value.get("returncode", value.get("second", {}).get("returncode")) != 0),
                   "returncode": value.get("returncode", value.get("second", {}).get("returncode"))}
            for name, value in negatives.items()
        },
        "negative_details": negatives,
        "scope_note": "The v4 tool does not consume or validate artifacts/manifest.json; this control binds the identity receipt and each listed XML member only, and makes no manifest-validation claim.",
    }
    if args.output.exists():
        raise AssertionError(f"refusing existing receipt {args.output}")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(args.output),
                      "output_bytes": args.output.stat().st_size,
                      "output_sha256": sha256(args.output),
                      "negative_count": len(negatives)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
