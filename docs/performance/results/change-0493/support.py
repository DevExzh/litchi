"""Shared custody helpers for the 0493 managed DOCX read-ahead evidence bundle.

The benchmark binaries are built and run by the coordinator.  These helpers
only provide deterministic paths, source manifests, receipt serialization and
content hashes; they do not invoke Cargo or execute a benchmark.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
TEMP = Path("/home/zhuhe/.cache/litchi-goal-0493")
TARGET_DIR = Path("/home/zhuhe/.cache/litchi-build-0493")
ENV = dict(
    os.environ,
    RUSTUP_TOOLCHAIN="1.98.1",
    CARGO_BUILD_JOBS="4",
    CARGO_INCREMENTAL="0",
    CARGO_PROFILE_RELEASE_DEBUG="1",
    RUSTFLAGS="-C force-frame-pointers=yes -C force-unwind-tables=yes",
    DEBUGINFOD_URLS="",
    LC_ALL="C",
    CARGO_TARGET_DIR=str(TARGET_DIR),
    PYTHONDONTWRITEBYTECODE="1",
)
ENV_KEYS = (
    "RUSTUP_TOOLCHAIN",
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_PROFILE_RELEASE_DEBUG",
    "RUSTFLAGS",
    "DEBUGINFOD_URLS",
    "LC_ALL",
    "CARGO_TARGET_DIR",
)

# The retained 0487 build records were produced before this lane moved Cargo
# output out of the repository.  Keep their exact receipt environment as a
# separate immutable baseline while all new gates and captures use ENV.
BASELINE_ENV_KEYS = tuple(key for key in ENV_KEYS if key != "CARGO_TARGET_DIR")


def baseline_environment() -> dict[str, str]:
    return {key: ENV[key] for key in BASELINE_ENV_KEYS}

# Every executable evidence helper is content-addressed in the frozen
# protocol.  This keeps the analyzer/verifier version coupled to the receipts
# that they interpret, while the gate's historical copies remain available
# for developmental attempts made before a freeze.
DRIVER_SCRIPTS = ("support.py", "gate.py")


def sha(path: Path | str) -> str:
    """Return the SHA-256 of one regular file."""

    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def read(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path: Path | str, value: Any) -> None:
    """Create one JSON receipt without replacing an existing artifact."""

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")


def meta(path: Path | str) -> dict[str, int | str]:
    path = Path(path)
    return {"bytes": path.stat().st_size, "sha256": sha(path)}


def source_names() -> list[str]:
    """List source inputs tracked or intentionally unignored in this checkout.

    Rust/TOML/lock sources are the compilation inputs.  XML templates under a
    crate source tree are also inputs to the DOCX and other OOXML owners.  The
    query deliberately excludes ignored build outputs and all documentation
    receipts.
    """

    output = subprocess.check_output(
        ["git", "ls-files", "-co", "--exclude-standard", "-z"],
        cwd=REPO,
    )
    names = output.decode("utf-8").split("\0")
    selected: list[str] = []
    for name in sorted(set(names)):
        if not name:
            continue
        path = Path(name)
        is_source = path.suffix in {".rs", ".toml", ".lock"}
        is_template = (
            name.startswith("crates/")
            and "/src/" in name
            and path.suffix == ".xml"
        )
        if (is_source or is_template) and (REPO / path).is_file():
            selected.append(name)
    return selected


def snapshot() -> dict[str, Any]:
    """Write or reuse a content-addressed source manifest.

    A manifest is never replaced.  This lets failed/development attempts keep
    their exact source identity while the final protocol can bind a later
    source revision explicitly.
    """

    sources = {name: sha(REPO / name) for name in source_names()}
    encoded = (json.dumps(sources, indent=2, sort_keys=True) + "\n").encode()
    digest = hashlib.sha256(encoded).hexdigest()
    directory = ROOT / "validation-sources"
    directory.mkdir(exist_ok=True)
    path = directory / f"{digest}.json"
    if path.exists():
        if sha(path) != digest:
            raise RuntimeError(f"existing source manifest has wrong digest: {path}")
    else:
        with path.open("xb") as stream:
            stream.write(encoded)
    return {
        "path": path.relative_to(ROOT).as_posix(),
        "sha256": digest,
        "files": len(sources),
    }


def environment() -> dict[str, str]:
    return {key: ENV[key] for key in ENV_KEYS}
