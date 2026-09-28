"""0810 source and artifact custody helpers.

This module only records identities and provides paths for the root-owned
drivers.  It does not execute a workload, invoke Cargo, or interpret results.
"""

from pathlib import Path
import hashlib
import json
import subprocess


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0810")
ROOT_INPUTS = {
    "Cargo.lock": P / "inputs/root-Cargo.lock",
    "rustfmt.toml": P / "inputs/rustfmt.toml",
}


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str):
    return json.loads(Path(path).read_text())


def write(path: Path | str, value) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def artifact(path: Path | str) -> dict[str, object]:
    path = Path(path)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def root_input_hashes() -> dict[str, str]:
    return {name: sha(path) for name, path in ROOT_INPUTS.items()}


def assert_root_inputs() -> dict[str, str]:
    frozen = root_input_hashes()
    for name, copy in ROOT_INPUTS.items():
        source_path = ROOT / name
        assert source_path.is_file(), f"missing root input {source_path}"
        assert sha(source_path) == frozen[name], f"root input changed: {name}"
    return frozen


def tracked_source_names() -> list[str]:
    raw = subprocess.check_output(
        [
            "git",
            "ls-files",
            "-z",
            "--",
            "crates",
            "Cargo.toml",
            "clippy.toml",
            ".cargo/config.toml",
            "rust-toolchain.toml",
        ],
        cwd=ROOT,
    )
    return [name for name in raw.decode().split("\0") if name]


def source() -> dict[str, object]:
    names = tracked_source_names()
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def changed_files(before: dict[str, object], after: dict[str, object]) -> set[str]:
    before_files = before["files"]
    after_files = after["files"]
    assert isinstance(before_files, dict) and isinstance(after_files, dict)
    return {
        name
        for name in before_files.keys() | after_files.keys()
        if before_files.get(name) != after_files.get(name)
    }
