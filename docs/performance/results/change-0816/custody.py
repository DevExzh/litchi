"""0816 source and artifact custody helpers.

The root agent owns execution of Cargo, the standalone benchmark, and all
workload capture.  This module only identifies inputs and verifies that the
frozen production and tool sources remain unchanged while those drivers run.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TOOL = ROOT / "tools/perf-execution"
TARGET = Path("/home/zhuhe/code/litchi-target-0816")

ROOT_INPUTS = {
    "Cargo.lock": P / "inputs/root-Cargo.lock",
    "rustfmt.toml": P / "inputs/rustfmt.toml",
}
TOOL_LOCK = P / "inputs/tool-Cargo.lock"

TOOL_FILES = {
    "tools/perf-execution/Cargo.lock",
    "tools/perf-execution/Cargo.toml",
    "tools/perf-execution/README.md",
    "tools/perf-execution/src/main.rs",
}
DRIVERS = ("custody.py", "build.py", "capture.py", "quality.py")
FROZEN_PACKET_FILES = (
    "plan.json",
    *DRIVERS,
    "architecture-inputs.json",
    "host.json",
    "origin.json",
    "toolchain.json",
    "lock-parity.json",
    "cgroup-limits.json",
    "root-inputs.json",
    "inputs/root-Cargo.lock",
    "inputs/rustfmt.toml",
    "inputs/tool-Cargo.lock",
    "protocol-review.md",
    "source-review.md",
)
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md":
        "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md":
        "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json":
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
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
    assert path.is_file() and not path.is_symlink(), f"missing artifact {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def root_input_hashes() -> dict[str, str]:
    return {name: sha(path) for name, path in ROOT_INPUTS.items()}


def assert_root_inputs() -> dict[str, str]:
    frozen = root_input_hashes()
    declared = read(P / "root-inputs.json")
    assert declared == frozen, "root input manifest differs from packet copies"
    for name, copy in ROOT_INPUTS.items():
        source = ROOT / name
        assert source.is_file(), f"missing root input {source}"
        assert sha(source) == frozen[name], f"root input changed: {name}"
        assert copy.is_file(), f"missing packet input {copy}"
    return frozen


def assert_tool_lock() -> dict[str, str]:
    assert TOOL_LOCK.is_file(), f"missing packet tool lock {TOOL_LOCK}"
    live = TOOL / "Cargo.lock"
    assert live.is_file(), f"missing tool lock {live}"
    frozen = {"sha256": sha(TOOL_LOCK), "bytes": TOOL_LOCK.stat().st_size}
    assert frozen["sha256"] == sha(live), "standalone tool lock differs from packet copy"
    assert frozen["bytes"] == live.stat().st_size, "standalone tool lock size changed"
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


def tool_source() -> dict[str, str]:
    actual = {
        str(path.relative_to(ROOT)): sha(path)
        for path in sorted(TOOL.rglob("*"))
        if path.is_file() and not path.is_symlink() and "target" not in path.parts
    }
    assert set(actual) == TOOL_FILES, (
        f"standalone tool file set changed: expected {sorted(TOOL_FILES)}, "
        f"got {sorted(actual)}"
    )
    return actual


def unchanged(frozen: dict[str, object]) -> None:
    """Require both production and standalone-tool identities to be stable."""
    production = frozen.get("production", frozen)
    tool = frozen.get("tool")
    assert isinstance(production, dict), "production source witness is malformed"
    assert source() == production, "production source changed"
    if tool is not None:
        assert isinstance(tool, dict), "tool source witness is malformed"
        expected = tool.get("files", tool)
        assert tool_source() == expected, "standalone tool source changed"


def architecture_hashes() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert isinstance(declared, dict) and len(declared) == 35, (
        f"expected 35 architecture inputs, found {len(declared)}"
    )
    actual = {name: sha(ROOT / name) for name in declared}
    assert actual == declared, "architecture input changed after packet freeze"
    return actual


def host_hash() -> str:
    return sha(P / "host.json")


def assert_host(expected: str | None = None) -> str:
    actual = host_hash()
    if expected is not None:
        assert actual == expected, "host descriptor changed after freeze"
    return actual


def packet_hashes() -> dict[str, str]:
    result = {}
    for name in FROZEN_PACKET_FILES:
        path = P / name
        assert path.is_file() and not path.is_symlink(), f"missing frozen packet input {path}"
        result[name] = sha(path)
    return result


def driver_hashes() -> dict[str, str]:
    return {name: sha(P / name) for name in DRIVERS}


def assert_unrelated() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    assert actual == UNRELATED, "an unrelated workspace file changed"
    return actual
