"""Custody helpers for the isolated 0806 quality amendment preflight.

This packet compares the immutable original production archive with the
mechanical ``unchecked_attributes`` constructor reuse supplied by the 0806
amendment handoff.  The helpers do not build or execute anything themselves.
"""

from __future__ import annotations

import hashlib
import json
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[4]
TARGET = Path("/home/zhuhe/code/litchi-target-0806/amendment-preflight")

SOURCE_FILES = {
    "litchi-ole-common-xml_attributes.rs": "crates/litchi-ole-common/src/xml_attributes.rs",
    "litchi-opc-xml_attributes.rs": "crates/litchi-opc/src/xml_attributes.rs",
    "litchi-opc-xml_attributes-tests.rs": "crates/litchi-opc/src/xml_attributes/tests.rs",
    "litchi-sign-xml_attributes.rs": "crates/litchi-sign/src/xml_attributes.rs",
    "litchi-xldm-xml_attributes.rs": "crates/litchi-xldm/src/xml_attributes.rs",
    "xml-minifier-xml_attributes.rs": "crates/xml-minifier/src/xml_attributes.rs",
}
HELPER_NAMES = tuple(name for name in SOURCE_FILES if not name.endswith("-tests.rs"))


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path: Path | str, value: Any) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def artifact(path: Path | str) -> dict[str, Any]:
    target = Path(path)
    return {"path": str(target), "bytes": target.stat().st_size, "sha256": sha(target)}


def relative_artifact(path: Path | str) -> dict[str, Any]:
    target = Path(path)
    return {
        "path": str(target.relative_to(P)),
        "bytes": target.stat().st_size,
        "sha256": sha(target),
    }


def source() -> dict[str, Any]:
    names = subprocess.check_output(
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
    ).decode().split("\0")
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in names if name},
    }


def archive_manifest(leg: str) -> dict[str, Any]:
    if leg not in ("before", "after"):
        raise ValueError(leg)
    directory = P / "source" / leg
    return {
        name: relative_artifact(directory / name) for name in sorted(SOURCE_FILES)
    }


def probe_manifest() -> dict[str, str]:
    directory = P / "probe-src"
    return {
        str(path.relative_to(P)): sha(path)
        for path in sorted(directory.rglob("*"))
        if path.is_file() and path.name != "Cargo.toml"
    }


def assert_artifact(descriptor: dict[str, Any], label: str) -> Path:
    path = Path(descriptor["path"])
    if not path.is_file() or path.is_symlink():
        raise AssertionError(f"{label}: missing artifact {path}")
    if path.stat().st_size != descriptor["bytes"] or sha(path) != descriptor["sha256"]:
        raise AssertionError(f"{label}: artifact changed {path}")
    return path
