#!/usr/bin/env python3
"""Run the retained 0424 Heaptrack capture protocol for the candidate role.

This wrapper creates an isolated ``candidate-profile`` root with a protocol
that explicitly identifies the candidate role, while retaining the frozen
control revision as ``control_revision``.  The shared capture driver then
retains the same report, catalog, raw ``time-v``/Heaptrack, verifier, and
per-corpus ``capture.json`` schema used by the historical control profile.
The output root is immutable and this driver never rewrites control evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
AUTHORIZED_PROFILE_FIELDS = frozenset({
    "role",
    "profile_role",
    "profile_source_protocol_sha256",
})


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON constant: {value}")


def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def strict_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicate_keys,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, ValueError, json.JSONDecodeError) as exc:
        raise SystemExit(f"cannot load JSON {path}: {exc}") from exc


def frozen_parent_protocol(protocol_path: Path) -> tuple[dict[str, Any], str]:
    """Load the root protocol and reject a substituted profile parent.

    ``--protocol`` remains a convenience for callers, but it must name a
    byte-identical copy of the frozen root protocol.  The candidate profile
    then records this hash in its authorized role metadata.
    """
    frozen_path = ROOT / "protocol.json"
    if not frozen_path.is_file():
        raise SystemExit(f"frozen root protocol is missing: {frozen_path}")
    frozen_bytes = frozen_path.read_bytes()
    supplied_bytes = protocol_path.read_bytes()
    frozen_hash = hashlib.sha256(frozen_bytes).hexdigest()
    if supplied_bytes != frozen_bytes:
        raise SystemExit(
            "candidate profile protocol must be byte-identical to the frozen root protocol"
        )
    parent = strict_json(frozen_path)
    if not isinstance(parent, dict):
        raise SystemExit("frozen root protocol must be a JSON object")
    if AUTHORIZED_PROFILE_FIELDS.intersection(parent):
        raise SystemExit(
            "frozen root protocol already contains candidate profile metadata"
        )
    return parent, frozen_hash


def candidate_protocol_for(
    parent: dict[str, Any], parent_hash: str
) -> dict[str, Any]:
    if parent.get("change") != 424:
        raise SystemExit("profile protocol must be the frozen 0424 protocol")
    if not parent.get("control_revision"):
        raise SystemExit("profile protocol must retain control_revision")
    for field in ("common_flags", "scope"):
        if field not in parent:
            raise SystemExit(f"frozen root protocol is missing required {field}")
    candidate = dict(parent)
    candidate.update({
        "role": "candidate",
        "profile_role": "candidate",
        "profile_source_protocol_sha256": parent_hash,
    })
    if set(candidate) != set(parent) | AUTHORIZED_PROFILE_FIELDS:
        raise SystemExit("candidate protocol changed fields outside authorized role metadata")
    for key, value in parent.items():
        if candidate.get(key) != value:
            raise SystemExit(f"candidate protocol changed frozen field: {key}")
    if candidate["common_flags"] != parent["common_flags"]:
        raise SystemExit("candidate protocol changed frozen common_flags")
    if candidate["scope"] != parent["scope"]:
        raise SystemExit("candidate protocol changed frozen scope")
    return candidate


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-root", type=Path, default=ROOT / "candidate-profile",
        help="fresh candidate profile root (default: candidate-profile)",
    )
    parser.add_argument(
        "--build", type=Path, default=ROOT / "measurement-build-candidate.json",
        help="standard 0423-style candidate build identity",
    )
    parser.add_argument(
        "--protocol", type=Path, default=ROOT / "protocol.json",
        help="frozen one-sample Heaptrack profile protocol",
    )
    parser.add_argument(
        "--verifier", type=Path, default=ROOT / "pinned" / "verify-report.py",
    )
    parser.add_argument(
        "--repo-root", type=Path, default=ROOT / "pinned",
        help="pinned validator root used by verify-report.py",
    )
    args = parser.parse_args()

    output = args.output_root.expanduser().resolve()
    build = args.build.expanduser().resolve()
    protocol_path = args.protocol.expanduser().resolve()
    verifier = args.verifier.expanduser().resolve()
    repo_root = args.repo_root.expanduser().resolve()
    if output.exists() and any(output.iterdir()):
        raise SystemExit(f"refusing to overwrite non-empty candidate profile: {output}")
    if not build.is_file():
        raise SystemExit(f"candidate build identity is missing: {build}")
    if not protocol_path.is_file():
        raise SystemExit(f"profile protocol is missing: {protocol_path}")
    protocol, parent_hash = frozen_parent_protocol(protocol_path)
    if output.exists():
        protocol_output = output / "protocol.json"
        if protocol_output.exists():
            raise SystemExit(f"refusing to overwrite candidate protocol: {protocol_output}")
    else:
        output.mkdir(parents=True)
    candidate_protocol = candidate_protocol_for(protocol, parent_hash)
    (output / "protocol.json").write_text(
        json.dumps(candidate_protocol, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    command = [
        sys.executable, "-B", str(ROOT / "capture.py"), "--role", "candidate",
        "--output-root", str(output), "--build", str(build),
        "--protocol", str(output / "protocol.json"), "--verifier", str(verifier),
        "--repo-root", str(repo_root),
    ]
    result = subprocess.run(command, cwd=ROOT, check=False)
    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
