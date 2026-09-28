"""Seal the retained 0808 early-stop packet and audit staged or committed blobs."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
DOCS = ("0808-pptx-direct-event-handling.md", "BASELINE.md", "CRUD_COVERAGE.md",
        "GOAL_AUDIT.md", "HOTSPOTS.md", "REPORT.md")
TARGET = Path("/home/zhuhe/code/litchi-target-0808")


class SealError(RuntimeError):
    pass


def require(value: bool, message: str) -> None:
    if not value:
        raise SealError(message)


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha(path: Path) -> str:
    return digest_bytes(path.read_bytes())


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def current_source() -> dict[str, str]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", "crates", "Cargo.toml", "clippy.toml",
         ".cargo/config.toml", "rust-toolchain.toml"], cwd=ROOT)
    names = [name for name in raw.decode().split("\0") if name]
    return {name: sha(ROOT / name) for name in names}


def payloads() -> dict[str, str]:
    result = {}
    for path in sorted(P.rglob("*")):
        require(not path.is_symlink(), f"packet contains symlink: {path}")
        if path.is_file() and path.name != "seal.json" and "__pycache__" not in path.parts:
            result[str(path.relative_to(P))] = sha(path)
    require(not list(P.rglob("__pycache__")), "packet contains Python cache")
    return result


def production() -> dict[str, str]:
    before = read(P / "build-before/source.json")
    files = before.get("files")
    require(isinstance(files, dict), "before source manifest malformed")
    live = current_source()
    require(live == files, "restored production source differs from before build")
    return {name: live[name] for name, old in files.items() if live.get(name) != old}


def make_seal() -> dict[str, Any]:
    validation = read(P / "early-stop-validation.json")
    require(validation.get("schema") == "litchi.performance.0808.early-stop-validation.v1"
            and validation.get("status") == "stopped-before-after-build"
            and validation.get("performance_measured") is False,
            "early-stop validation is missing or not terminal")
    documents = {}
    for name in DOCS:
        path = ROOT / "docs/performance" / name
        require(path.is_file() and not path.is_symlink(), f"missing report document: {name}")
        documents[f"docs/performance/{name}"] = sha(path)
    value = {
        "schema": "litchi.performance.0808.early-stop-seal.v1",
        "status": "stopped-before-after-build",
        "production": production(),
        "files": payloads(),
        "documents": documents,
    }
    require(value["production"] == {}, "early-stop packet changed production source")
    return value


def tree_listing(index: bool) -> set[str]:
    command = (["git", "diff", "--cached", "--name-only", "-z"] if index else
               ["git", "diff-tree", "--no-commit-id", "--name-only", "-r", "-z", "HEAD"])
    return {name for name in subprocess.check_output(command, cwd=ROOT).decode().split("\0") if name}


def check_tree(seal: dict[str, Any], *, index: bool) -> None:
    files = seal["files"]
    expected = {str(P.relative_to(ROOT) / name): digest for name, digest in files.items()}
    expected.update(seal["documents"])
    encoded = json.dumps(seal, indent=2, sort_keys=True) + "\n"
    expected[str((P / "seal.json").relative_to(ROOT))] = digest_bytes(encoded.encode())
    changed = tree_listing(index)
    require(changed == set(expected),
            f"{'index' if index else 'HEAD'} paths differ: added={sorted(changed - set(expected))}, "
            f"missing={sorted(set(expected) - changed)}")
    revision = ":" if index else "HEAD:"
    process = subprocess.Popen(["git", "cat-file", "--batch"], cwd=ROOT,
                               stdin=subprocess.PIPE, stdout=subprocess.PIPE)
    assert process.stdin is not None and process.stdout is not None
    try:
        for name, expected_sha in expected.items():
            process.stdin.write((revision + name + "\n").encode())
            process.stdin.flush()
            header = process.stdout.readline().decode().split()
            require(len(header) == 3 and header[1] == "blob", f"{name}: staged blob missing")
            data = process.stdout.read(int(header[2]))
            require(process.stdout.read(1) == b"\n" and digest_bytes(data) == expected_sha,
                    f"{name}: staged blob changed")
    finally:
        process.stdin.close()
    require(process.wait() == 0, "git cat-file failed")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check-index", action="store_true")
    parser.add_argument("--check-head", action="store_true")
    args = parser.parse_args()
    require(args.write or args.check_index or args.check_head,
            "use --write, --check-index, or --check-head")
    value = make_seal()
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    path = P / "seal.json"
    if args.write:
        path.write_text(encoded, encoding="utf-8")
    require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
            "seal.json is stale")
    if args.check_index:
        check_tree(value, index=True)
    if args.check_head:
        check_tree(value, index=False)
    print(f"0808 early-stop seal PASS: {len(value['files'])} payloads + seal + "
          f"{len(value['documents'])} documents")


if __name__ == "__main__":
    try:
        main()
    except (OSError, SealError, subprocess.CalledProcessError) as error:
        print(f"early-stop seal failed: {error}")
        raise SystemExit(1)
