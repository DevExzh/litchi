#!/usr/bin/env python3
"""Prepare one immutable DrawingML raw baseline/candidate source pair.

Each raw lane has its own production parent.  The candidate carries the
reviewed protocol adapter; this helper copies only that adapter into a clean
detached baseline and records the exact dirty-path allowlist.  It does not
run Cargo, a native probe, or a timing command.  A later paired runner must
revalidate this receipt before it starts either build or measurement.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, Mapping


SCHEMA = "litchi.drawingml.raw-preparation.v1"
# Version one is a frozen historical protocol.  Its consumer adapter is
# tested against this approved standalone source capture, never against the
# current repository's lib.rs (which may have moved on to another API role).
FROZEN_CONSUMER_API_SOURCE_REF = "805c08facf9ff7c92c7f05b66bde1e2bdb8524f2"
FROZEN_CONSUMER_API_SOURCE_PATH = "tools/perf-baseline/src/lib.rs"
FROZEN_CONSUMER_API_SOURCE_BLOB_SHA1 = "2ec7f936b22bef7468983134b0b5f05386f7f8a8"
FROZEN_CONSUMER_API_SOURCE_SHA256 = "4acd072fd68c094e24524862194da9676260e1fb25f6ea3dbb539197d2838b5b"
FROZEN_CONSUMER_API_SOURCE_BYTES = 3_088_209
HISTORICAL_WORKSPACE_LOCK_SHA256 = "52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6"
HISTORICAL_WORKSPACE_LOCK_BYTES = 139591
# Kept as compatibility names for callers that only need the fixed historical
# baseline identity. Candidate roots use their separately supplied lock hash.
WORKSPACE_LOCK_SHA256 = HISTORICAL_WORKSPACE_LOCK_SHA256
WORKSPACE_LOCK_BYTES = HISTORICAL_WORKSPACE_LOCK_BYTES
ADAPTER_PATHS = (
    "Cargo.lock",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/drawingml_raw_probe.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_native.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_allocator.rs",
    "tools/perf-baseline/src/bin/support/drawingml_counting_allocator.rs",
    "tools/native_qualification/drawingml_raw_validator.py",
    "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
)

# The DrawingML probe is shared with historical RTF, XLS, DOCX, and PPTX
# benchmark parents. The historical baseline retains the former RTF/XLS
# consumer spellings, while the current DOCX/PPTX exports moved under their
# validation modules. The baseline adapter below applies exactly these six
# reviewed role-specific substitutions while leaving every DrawingML route
# byte-identical. Candidate sources are never transformed.
CONSUMER_API_ADAPTER_PATH = "tools/perf-baseline/src/lib.rs"
CONSUMER_API_REPLACEMENTS = (
    (
        b"litchi_rtf::streaming::Limits::new",
        b"litchi_rtf::write::StreamingRtfLimits::new",
    ),
    (
        b"litchi_rtf::streaming::Writer::new",
        b"litchi_rtf::write::StreamingRtfWriter::new",
    ),
    (
        b"litchi_xls::validation::Limits::default()",
        b"litchi_xls::validation::XlsValidationLimits::default()",
    ),
    (
        b"litchi_docx::validation::validate_read_at",
        b"litchi_docx::validate_read_at",
    ),
    (
        b"litchi_pptx::validation::validate_source",
        b"litchi_pptx::validate_source",
    ),
    (
        b"litchi_pptx::validation::Limits::default()",
        b"litchi_pptx::PptxValidationLimits::default()",
    ),
)
CONSUMER_API_TRANSFORM_COUNT = len(CONSUMER_API_REPLACEMENTS)
# This hash is the checked-in composed candidate identity. The historical
# baseline adapter below reverses these spellings only in the baseline copy.
CONSUMER_API_CURRENT_SHA256 = "4acd072fd68c094e24524862194da9676260e1fb25f6ea3dbb539197d2838b5b"
CONSUMER_API_HISTORICAL_SHA256 = "280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a"
CONSUMER_API_TRANSFORM_BYTES = sum(
    len(historical) - len(current)
    for current, historical in CONSUMER_API_REPLACEMENTS
)

# Filled from the reviewed adapter in this leaf.  Keeping this table in the
# preparer means a future protocol edit cannot silently become historical
# baseline evidence merely because it is copied into a worktree.
STATIC_REVIEWED_ADAPTER_SHA256 = {
    # The current shared consumer manifest includes the DOCX and ODF probe
    # registrations composed after the original DrawingML adapter review.
    # Keep the reviewed identity bound to the actual checked-in candidate;
    # the detached historical RTF manifest pins remain separate.
    "tools/perf-baseline/Cargo.toml": "f18a17d5ca2bb0848e73c53ace39550ba38dbd7a75fab76c5a0b0af4f435984a",
    "tools/perf-baseline/src/lib.rs": CONSUMER_API_CURRENT_SHA256,
    "tools/perf-baseline/src/drawingml_raw_probe.rs": "539e4641b1eecda7ef97f1bfca2e526348f0a0d4941cdec781cadb3c99a18477",
    "tools/perf-baseline/src/bin/drawingml_raw_native.rs": "c64e5cb9436622d383a55f6720c940344c7575cc56404ab6459133afbc63116e",
    "tools/perf-baseline/src/bin/drawingml_raw_allocator.rs": "6dd60e14b760bf7562017ddc99b9356699a13155eb88a46cdb62a81aa3b0d3b9",
    "tools/perf-baseline/src/bin/support/drawingml_counting_allocator.rs": "3db5f3a49cf1c3e5ef496f6f3a787f94e74fa3d995879af578a60cb910d13117",
    "tools/native_qualification/drawingml_raw_validator.py": "1793863f8eecbcadcc1ec071547ffc9188b960704b3fd5be2cd07d55c2664b15",
    "tools/native_qualification/fixtures/drawingml_raw_fixtures.json": "e72237ec2095535667fa34e962e37950c96ef558eef151839a4e8235a1256925",
}


LOCK_PATHS = frozenset(("Cargo.lock", "tools/perf-baseline/Cargo.lock"))


def _reviewed_adapter_hashes(
    perf_lock_sha256: str,
    workspace_lock_sha256: str,
    role: str = "candidate",
) -> dict[str, str]:
    if role not in {"baseline", "candidate"}:
        raise ValueError(f"unknown DrawingML adapter role: {role}")
    result = {
        **STATIC_REVIEWED_ADAPTER_SHA256,
        "Cargo.lock": workspace_lock_sha256,
        "tools/perf-baseline/Cargo.lock": perf_lock_sha256,
    }
    if role == "baseline":
        result[CONSUMER_API_ADAPTER_PATH] = CONSUMER_API_HISTORICAL_SHA256
    return result


class PreparationError(ValueError):
    """A source preparation or receipt invariant failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _consumer_api_adapter_receipt() -> dict[str, Any]:
    return {
        "path": CONSUMER_API_ADAPTER_PATH,
        "candidate_sha256": CONSUMER_API_CURRENT_SHA256,
        "baseline_sha256": CONSUMER_API_HISTORICAL_SHA256,
        "replacement_count": CONSUMER_API_TRANSFORM_COUNT,
        "changed_bytes": CONSUMER_API_TRANSFORM_BYTES,
        "replacements": [
            {
                "candidate": current.decode("ascii"),
                "baseline": historical.decode("ascii"),
                "count": 1,
            }
            for current, historical in CONSUMER_API_REPLACEMENTS
        ],
    }


def _adapt_baseline_consumer_adapter(relative: str, source: bytes) -> bytes:
    """Apply only the reviewed historical consumer spellings to a baseline."""

    if relative != CONSUMER_API_ADAPTER_PATH:
        return source
    _expect(
        _sha256_bytes(source) == CONSUMER_API_CURRENT_SHA256,
        "candidate consumer adapter is not the reviewed composed source",
    )
    adapted = source
    replacement_count = 0
    changed_bytes = 0
    for current, historical in CONSUMER_API_REPLACEMENTS:
        current_count = adapted.count(current)
        historical_count = adapted.count(historical)
        _expect(
            current_count == 1 and historical_count == 0,
            "RTF consumer adapter spelling/count changed",
        )
        adapted = adapted.replace(current, historical)
        replacement_count += current_count
        changed_bytes += len(historical) - len(current)
    _expect(
        replacement_count == CONSUMER_API_TRANSFORM_COUNT,
        "consumer adapter replacement count changed",
    )
    _expect(
        changed_bytes == CONSUMER_API_TRANSFORM_BYTES,
        "consumer adapter byte delta changed",
    )
    _expect(
        _sha256_bytes(adapted) == CONSUMER_API_HISTORICAL_SHA256,
        "historical consumer adapter hash changed",
    )
    return adapted


def _sha256_file(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"identity path is not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _git(root: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(arguments)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _git_binary(root: Path, *arguments: str) -> bytes:
    result = subprocess.run(
        ["git", "-C", str(root), *arguments],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise PreparationError(f"git {' '.join(arguments)} failed in {root}: {detail}")
    return bytes(result.stdout)


def _root(value: Path) -> Path:
    root = value.resolve(strict=False)
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve(strict=False)
    _expect(root == top, f"source root is not a worktree top-level: {root}")
    return root


def _frozen_consumer_api_source(root: Path) -> tuple[bytes, dict[str, Any]]:
    """Read the approved v1 perflib blob without consulting the live file."""

    root = _root(root)
    actual_ref = _git(root, "rev-parse", "--verify", f"{FROZEN_CONSUMER_API_SOURCE_REF}^{{commit}}").strip()
    _expect(actual_ref == FROZEN_CONSUMER_API_SOURCE_REF, "v1 frozen consumer ref differs")
    payload = _git_binary(root, "show", f"{FROZEN_CONSUMER_API_SOURCE_REF}:{FROZEN_CONSUMER_API_SOURCE_PATH}")
    _expect(len(payload) == FROZEN_CONSUMER_API_SOURCE_BYTES, "v1 frozen consumer byte count differs")
    _expect(_sha256_bytes(payload) == FROZEN_CONSUMER_API_SOURCE_SHA256, "v1 frozen consumer SHA-256 differs")
    blob = hashlib.sha1(f"blob {len(payload)}\0".encode("ascii") + payload).hexdigest()
    _expect(blob == FROZEN_CONSUMER_API_SOURCE_BLOB_SHA1, "v1 frozen consumer Git blob differs")
    return payload, {
        "ref": FROZEN_CONSUMER_API_SOURCE_REF,
        "path": FROZEN_CONSUMER_API_SOURCE_PATH,
        "blob_sha1": FROZEN_CONSUMER_API_SOURCE_BLOB_SHA1,
        "sha256": FROZEN_CONSUMER_API_SOURCE_SHA256,
        "bytes": FROZEN_CONSUMER_API_SOURCE_BYTES,
    }


def _status_paths(status: str) -> tuple[str, ...]:
    paths: list[str] = []
    for line in status.splitlines():
        _expect(len(line) >= 4, f"malformed Git status line: {line!r}")
        value = line[3:]
        _expect(" -> " not in value, f"renames are not allowed in a preparation adapter: {line!r}")
        paths.append(value)
    return tuple(sorted(paths))


def _identity(root: Path) -> dict[str, Any]:
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    head = _git(root, "rev-parse", "HEAD").strip()
    hashes = {path: _sha256_file(root / path) for path in ADAPTER_PATHS}
    return {
        "root": str(root),
        "head": head,
        "tree": tree,
        "status_paths": list(_status_paths(status)),
        "status_sha256": _sha256_bytes(status.encode()),
        "diff_sha256": _sha256_bytes(diff.encode()),
        "source_tree_sha256": _sha256_bytes((tree + "\0" + status + "\0" + diff).encode()),
        "adapter_hashes": {
            path: {"sha256": digest, "bytes": size} for path, (digest, size) in hashes.items()
        },
        "workspace_lock_sha256": hashes["Cargo.lock"][0],
        "workspace_lock_bytes": hashes["Cargo.lock"][1],
        "perf_lock_sha256": hashes["tools/perf-baseline/Cargo.lock"][0],
        "perf_lock_bytes": hashes["tools/perf-baseline/Cargo.lock"][1],
    }


def _validate_adapter(
    identity: Mapping[str, Any],
    label: str,
    *,
    reviewed: bool,
    role: str = "candidate",
    perf_lock_sha256: str,
    workspace_lock_sha256: str,
    workspace_lock_bytes: int,
) -> None:
    statuses = identity.get("status_paths")
    _expect(isinstance(statuses, list), f"{label} status paths are missing")
    _expect(set(statuses).issubset(set(ADAPTER_PATHS)), f"{label} changed outside adapter allowlist: {statuses!r}")
    hashes = identity.get("adapter_hashes")
    _expect(isinstance(hashes, dict) and set(hashes) == set(ADAPTER_PATHS), f"{label} adapter hashes are incomplete")
    _expect(identity.get("workspace_lock_sha256") == workspace_lock_sha256, f"{label} workspace lock differs")
    _expect(identity.get("workspace_lock_bytes") == workspace_lock_bytes, f"{label} workspace lock byte count differs")
    _expect(identity.get("perf_lock_sha256") == perf_lock_sha256, f"{label} does not use the selected perf-baseline lock")
    if reviewed:
        for path, expected in _reviewed_adapter_hashes(
            perf_lock_sha256, workspace_lock_sha256, role
        ).items():
            observed = hashes[path].get("sha256")
            _expect(observed == expected, f"{label} reviewed adapter hash mismatch: {path}")


def _validate_receipt_data(
    receipt: Mapping[str, Any],
    *,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    perf_lock_sha256: str,
    candidate_workspace_lock_sha256: str,
    candidate_workspace_lock_bytes: int,
) -> dict[str, Any]:
    _expect(receipt.get("schema") == SCHEMA, "preparation receipt schema mismatch")
    _expect(receipt.get("baseline_revision") == baseline_revision, "receipt baseline revision differs")
    _expect(receipt.get("candidate_revision") == candidate_revision, "receipt candidate revision differs")
    _expect(receipt.get("adapter_paths") == list(ADAPTER_PATHS), "adapter path allowlist changed")
    _expect(
        receipt.get("baseline_workspace_lock_sha256") == HISTORICAL_WORKSPACE_LOCK_SHA256,
        "receipt baseline workspace lock is not the historic 52ed lock",
    )
    _expect(
        receipt.get("candidate_workspace_lock_sha256") == candidate_workspace_lock_sha256,
        "receipt candidate workspace lock differs",
    )
    _expect(
        receipt.get("baseline_workspace_lock_bytes") == HISTORICAL_WORKSPACE_LOCK_BYTES,
        "receipt baseline workspace lock byte count differs",
    )
    _expect(
        receipt.get("candidate_workspace_lock_bytes") == candidate_workspace_lock_bytes,
        "receipt candidate workspace lock byte count differs",
    )
    _expect(receipt.get("perf_lock_sha256") == perf_lock_sha256, "receipt perf-baseline lock differs")
    _expect(
        receipt.get("role_consumer_adapter") == _consumer_api_adapter_receipt(),
        "role-aware consumer adapter receipt changed",
    )
    _expect(
        receipt.get("reviewed_adapter_sha256") == {
            "baseline": _reviewed_adapter_hashes(
                perf_lock_sha256, HISTORICAL_WORKSPACE_LOCK_SHA256, "baseline"
            ),
            "candidate": _reviewed_adapter_hashes(
                perf_lock_sha256, candidate_workspace_lock_sha256, "candidate"
            ),
        },
        "reviewed adapter hashes changed",
    )
    baseline = receipt.get("baseline")
    candidate = receipt.get("candidate")
    _expect(isinstance(baseline, dict) and isinstance(candidate, dict), "source identities are missing")
    _expect(Path(str(baseline.get("root"))).resolve(strict=False) == baseline_root, "baseline root differs")
    _expect(Path(str(candidate.get("root"))).resolve(strict=False) == candidate_root, "candidate root differs")
    _expect(baseline.get("head") == baseline_revision, "baseline HEAD differs")
    _expect(candidate.get("head") == candidate_revision, "candidate HEAD differs")
    _expect(receipt.get("perf_lock_bytes") == candidate.get("perf_lock_bytes"), "receipt perf-baseline lock byte count differs")
    _validate_adapter(
        baseline,
        "baseline",
        reviewed=True,
        role="baseline",
        perf_lock_sha256=perf_lock_sha256,
        workspace_lock_sha256=HISTORICAL_WORKSPACE_LOCK_SHA256,
        workspace_lock_bytes=HISTORICAL_WORKSPACE_LOCK_BYTES,
    )
    _validate_adapter(
        candidate,
        "candidate",
        reviewed=True,
        role="candidate",
        perf_lock_sha256=perf_lock_sha256,
        workspace_lock_sha256=candidate_workspace_lock_sha256,
        workspace_lock_bytes=candidate["workspace_lock_bytes"],
    )
    _expect(candidate.get("status_paths") == [], "candidate source must be clean")
    _expect(
        {
            path: baseline["adapter_hashes"][path]
            for path in ADAPTER_PATHS
            if path not in LOCK_PATHS and path != CONSUMER_API_ADAPTER_PATH
        }
        == {
            path: candidate["adapter_hashes"][path]
            for path in ADAPTER_PATHS
            if path not in LOCK_PATHS and path != CONSUMER_API_ADAPTER_PATH
        },
        "non-lock adapter bytes differ",
    )
    _expect(receipt.get("baseline_status_paths") == baseline.get("status_paths"), "baseline status receipt differs")
    return dict(receipt)


def validate_preparation_receipt(
    receipt_path: Path,
    *,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    perf_lock_sha256: str,
    candidate_workspace_lock_sha256: str,
    candidate_workspace_lock_bytes: int,
) -> dict[str, Any]:
    receipt_path = receipt_path.resolve(strict=False)
    _expect(not receipt_path.is_symlink() and receipt_path.is_file(), f"receipt is missing: {receipt_path}")
    try:
        receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read preparation receipt: {error}") from error
    _expect(isinstance(receipt, dict), "receipt root is not an object")
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    expected = _validate_receipt_data(
        receipt,
        baseline_root=baseline_root,
        candidate_root=candidate_root,
        baseline_revision=baseline_revision,
        candidate_revision=candidate_revision,
        perf_lock_sha256=perf_lock_sha256,
        candidate_workspace_lock_sha256=candidate_workspace_lock_sha256,
        candidate_workspace_lock_bytes=candidate_workspace_lock_bytes,
    )
    for label, current in (("baseline", _identity(baseline_root)), ("candidate", _identity(candidate_root))):
        recorded = expected[label]
        for key in (
            "root",
            "head",
            "tree",
            "status_paths",
            "status_sha256",
            "diff_sha256",
            "source_tree_sha256",
            "adapter_hashes",
            "workspace_lock_sha256",
            "workspace_lock_bytes",
            "perf_lock_sha256",
            "perf_lock_bytes",
        ):
            _expect(current.get(key) == recorded.get(key), f"{label} identity changed: {key}")
    return expected


def prepare(
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    receipt_path: Path,
    perf_lock_sha256: str,
    candidate_workspace_lock_sha256: str,
    candidate_workspace_lock_bytes: int,
) -> None:
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    _expect(baseline_root != candidate_root, "baseline and candidate roots must differ")
    _expect(_git(baseline_root, "rev-parse", "HEAD").strip() == baseline_revision, "baseline HEAD differs")
    _expect(_git(candidate_root, "rev-parse", "HEAD").strip() == candidate_revision, "candidate HEAD differs")
    _expect(not _git(candidate_root, "status", "--porcelain=v1", "--untracked-files=all"), "candidate source must be clean")
    before_status = _status_paths(_git(baseline_root, "status", "--porcelain=v1", "--untracked-files=all"))
    _expect(set(before_status).issubset(set(ADAPTER_PATHS)), "baseline already contains an unapproved change")
    candidate = _identity(candidate_root)
    _expect(
        candidate["workspace_lock_bytes"] == candidate_workspace_lock_bytes,
        "candidate workspace lock byte count differs",
    )
    _validate_adapter(
        candidate,
        "candidate",
        reviewed=True,
        role="candidate",
        perf_lock_sha256=perf_lock_sha256,
        workspace_lock_sha256=candidate_workspace_lock_sha256,
        workspace_lock_bytes=candidate["workspace_lock_bytes"],
    )
    for relative in ADAPTER_PATHS:
        if relative in LOCK_PATHS:
            continue
        source = candidate_root / relative
        target = baseline_root / relative
        _expect(not source.is_symlink() and source.is_file(), f"candidate adapter path is missing: {source}")
        _expect(not target.is_symlink(), f"baseline adapter target is a symlink: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        if relative == CONSUMER_API_ADAPTER_PATH:
            target.write_bytes(_adapt_baseline_consumer_adapter(relative, source.read_bytes()))
        else:
            shutil.copyfile(source, target)
    baseline = _identity(baseline_root)
    _validate_adapter(
        baseline,
        "baseline",
        reviewed=True,
        role="baseline",
        perf_lock_sha256=perf_lock_sha256,
        workspace_lock_sha256=HISTORICAL_WORKSPACE_LOCK_SHA256,
        workspace_lock_bytes=HISTORICAL_WORKSPACE_LOCK_BYTES,
    )
    _expect(baseline.get("head") == baseline_revision, "baseline HEAD changed during preparation")
    _expect(candidate.get("head") == candidate_revision, "candidate HEAD changed during preparation")
    _expect(
        {
            path: baseline["adapter_hashes"][path]
            for path in ADAPTER_PATHS
            if path not in LOCK_PATHS and path != CONSUMER_API_ADAPTER_PATH
        }
        == {
            path: candidate["adapter_hashes"][path]
            for path in ADAPTER_PATHS
            if path not in LOCK_PATHS and path != CONSUMER_API_ADAPTER_PATH
        },
        "non-lock adapter bytes differ",
    )
    receipt = {
        "schema": SCHEMA,
        "baseline_revision": baseline_revision,
        "candidate_revision": candidate_revision,
        "adapter_paths": list(ADAPTER_PATHS),
        "baseline_status_paths": list(baseline.get("status_paths", [])),
        "baseline_workspace_lock_sha256": HISTORICAL_WORKSPACE_LOCK_SHA256,
        "baseline_workspace_lock_bytes": HISTORICAL_WORKSPACE_LOCK_BYTES,
        "candidate_workspace_lock_sha256": candidate_workspace_lock_sha256,
        "candidate_workspace_lock_bytes": candidate["workspace_lock_bytes"],
        "perf_lock_sha256": perf_lock_sha256,
        "perf_lock_bytes": candidate["perf_lock_bytes"],
        "role_consumer_adapter": _consumer_api_adapter_receipt(),
        "reviewed_adapter_sha256": {
            "baseline": _reviewed_adapter_hashes(
                perf_lock_sha256, HISTORICAL_WORKSPACE_LOCK_SHA256, "baseline"
            ),
            "candidate": _reviewed_adapter_hashes(
                perf_lock_sha256, candidate_workspace_lock_sha256, "candidate"
            ),
        },
        "baseline": baseline,
        "candidate": candidate,
    }
    receipt_path = receipt_path.resolve(strict=False)
    _expect(not receipt_path.is_symlink() and not receipt_path.exists(), f"refusing to overwrite receipt: {receipt_path}")
    receipt_path.parent.mkdir(parents=True, exist_ok=True)
    with receipt_path.open("x", encoding="utf-8") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write("\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--perf-lock-sha256", required=True)
    parser.add_argument("--candidate-workspace-lock-sha256", required=True)
    parser.add_argument("--candidate-workspace-lock-bytes", required=True, type=int)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        prepare(
            arguments.baseline_root,
            arguments.candidate_root,
            arguments.baseline_revision,
            arguments.candidate_revision,
            arguments.output,
            arguments.perf_lock_sha256,
            arguments.candidate_workspace_lock_sha256,
            arguments.candidate_workspace_lock_bytes,
        )
    except (OSError, PreparationError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
