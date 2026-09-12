"""Pure gate and receipt checks used by the PPTX InkAction operator."""

from __future__ import annotations

import re
from pathlib import Path


_SHA256 = re.compile(r"[0-9a-fA-F]{64}\Z")


def output_path_without_symlink(value: str) -> Path:
    """Canonicalize an output path while rejecting symlinked lineage."""

    path = Path(value)
    if not path.is_absolute():
        path = Path.cwd() / path
    current = Path(path.anchor)
    for component in path.parts[1:]:
        if component in ("", "."):
            continue
        if component == "..":
            current = current.parent
            continue
        current /= component
        if current.is_symlink():
            raise ValueError(f"output path lineage contains a symlink: {current}")
    return path.resolve(strict=False)


def receipt_digest(receipt: str) -> str:
    """Return the digest field from a ``sha256sum`` receipt.

    The path field is deliberately ignored.  A malformed or missing digest is
    an operator error rather than a binary-change result.
    """

    fields = receipt.split()
    if len(fields) < 2 or _SHA256.fullmatch(fields[0]) is None:
        raise ValueError("binary hash receipt does not contain a SHA-256 digest and path")
    return fields[0].lower()


def binary_digests_match(before_receipt: str, after_digest: str) -> bool:
    """Compare the before-receipt digest with a separately computed digest."""

    if _SHA256.fullmatch(after_digest) is None:
        raise ValueError("computed binary digest is not SHA-256")
    return receipt_digest(before_receipt) == after_digest.lower()


def allowed_scenario_stages(
    host_probe_exit: int | None,
    matrix_exit: int | None,
) -> tuple[str, ...]:
    """Return the downstream stages permitted by completed prerequisites."""

    stages = ["host-probe"]
    if host_probe_exit != 0:
        return tuple(stages)
    stages.append("matrix-correctness")
    if matrix_exit != 0:
        return tuple(stages)
    stages.append("lanes")
    return tuple(stages)
