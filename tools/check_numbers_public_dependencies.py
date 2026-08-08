#!/usr/bin/env python3
"""Reject litchi-numbers API references to its private wire dependency.

Cargo's public-dependency manifest metadata is unstable.  Keep the workspace
manifests usable by stable Cargo and mark only `litchi_numbers_wire` private on
the final rustc invocation instead.  `RUSTC_BOOTSTRAP` enables that rustc-only
extern modifier on the stable toolchain used by CI.

Shared semantic values from `litchi_iwa_common` remain approved: this command
does not mark that dependency private.
"""

from __future__ import annotations

import os
import subprocess
import sys
from collections.abc import Mapping
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def command() -> tuple[str, ...]:
    """Return the deterministic compiler invocation used by the gate."""
    return (
        "cargo",
        "rustc",
        "--locked",
        "--package",
        "litchi-numbers",
        "--lib",
        "--",
        "-Zunstable-options",
        "--extern",
        "priv:litchi_numbers_wire",
        "-Dexported-private-dependencies",
    )


def environment(source: Mapping[str, str] | None = None) -> dict[str, str]:
    """Build the stable-runner environment without mutating the caller's map."""
    result = dict(os.environ if source is None else source)
    result["RUSTC_BOOTSTRAP"] = "1"
    return result


def main() -> int:
    completed = subprocess.run(
        command(),
        cwd=ROOT,
        env=environment(),
        check=False,
    )
    return completed.returncode


if __name__ == "__main__":
    sys.exit(main())
