#!/usr/bin/env python3
"""Validate the reusable RTF and OLE/XLS native probe receipts.

The Rust examples own the operation implementation, while this validator
checks the emitted route contract as well as the paired-run boundary: exact
baseline identity, reviewed preparation/manifest/runner provenance,
binary/lock/probe/fixture hashes, fixed ABBA order, quiet-process/cap receipts,
and complete route output streams.  It never interprets elapsed values as an
acceptance result.  Allocation instrumentation is intentionally kept out of
the normal timing binaries; its separate, non-claimable phase is recorded by
the protocol document.
"""

from __future__ import annotations

import csv
import hashlib
import json
import math
import os
from pathlib import Path
import re
import sys
from typing import Any, Iterable


SCHEMA = "litchi.native-qualification.v1"
SUMMARY_SCHEMA = "litchi.legacy-native-qualification.v1"
ABBA_ORDER = ("before", "after", "after", "before")
TARGET_PRE_FLIGHT = 19_200_000_000
TARGET_HARD = 20_000_000_000
MEMORY_HARD = 20_000_000_000
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
HASH = re.compile(r"^[0-9a-f]{64}$")
ROOT = Path(__file__).resolve().parents[2]
RUNNER_PATH = ROOT / "tools/native_qualification_runner.py"
PREPARER_PATH = ROOT / "tools/native_qualification/prepare_legacy_native_inputs.py"

REVIEWED_INPUTS: dict[str, dict[str, Any]] = {
    "legacy-rtf": {
        "probe_relative": "tools/perf-baseline/examples/refusal_contract_bench.rs",
        "lock_relative": "tools/perf-baseline/Cargo.lock",
        "probe_sha256": {
            "before": "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81",
            "after": "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81",
        },
        "lock_sha256": {
            "before": "57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f",
            "after": "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c",
        },
        "adapter": {
            "relative_path": "tools/perf-baseline/examples/refusal_contract_bench.rs",
            "probe_sha256": "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81",
            "manifest_relative_path": "tools/perf-baseline/Cargo.toml",
            "manifest_sha256": {
                "before": "af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676",
                "after": "61e103809753ffc528b2cd2c2e00a1525d40fc4e6360668ad982bec22a60bb73",
            },
            "features": {
                "before": "rtf-legacy-api",
                "after": "rtf-current-api",
            },
            "native_error_api": {
                "before": "litchi_rtf::edit::Error::{UnsupportedSource,ProtectedDocument}",
                "after": "litchi_rtf::edit::Error::Refused(Refusal)",
            },
            "apply_result_api": {
                "before": "litchi_rtf::Package::apply -> litchi_rtf::Package",
                "after": "litchi_rtf::Package::apply -> litchi_rtf::edit::Commit",
            },
            "live_current_probe": {
                "relative_path": "crates/litchi-rtf/examples/refusal_contract_bench.rs",
                "sha256": "1726acb58953ae3f57d9b74207f75678a3bcc7ac4c993443aac5583d3943f791",
                "lock_sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
                "lock_source": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/locks/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock",
                "api": "current litchi-rtf workspace example; not part of the baseline adapter source",
            },
            "historical_failed_compile": {
                "source_sha256": "70077fcec439fda2fdf9e9635f5db2f1f7c86ed31f3f972a95c5e5e3a607a585",
                "source_relative_path": "crates/litchi-rtf/examples/refusal_contract_bench.rs",
                "api_context": "current litchi-rtf consumer at candidate fence 35e37dbede6657dd023b114a1536862c18541946",
                "receipt": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/rtf-probe-compile/check-before.json",
                "receipt_sha256": "f3af926353fc3db4f74221f62ed4f4743cc6b31c359ab42415c8f1500e33bda4",
                "log": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/rtf-probe-compile/check-before.log",
                "log_sha256": "119d4216675ee79a7e89e17579f019dd4fe3a6d43c08a8bd22d87c37622fd103",
                "exit": 101,
            },
        },
        "fixture_relatives": (),
        "fixture_id": "rtf-refusal-contract-v1",
    },
    "legacy-ole-properties": {
        # This is the reviewed fdb139 probe, not the newer checkout copy.  The
        # baseline materialization selects the public owner API with the
        # legacy property API; the candidate adds only the reviewed CFB cap.
        "probe_relative": "tools/perf-baseline/examples/ole_w5_probe.rs",
        "lock_relative": "tools/perf-baseline/Cargo.lock",
        "probe_sha256": {
            "before": "6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed",
            "after": "e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782",
        },
        "adapter": {
            "source_fence": "fdb139197d08677d6fa0e14d536e798e380e4cb1",
            "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
            "baseline": {
                "source_fence": "fdb139197d08677d6fa0e14d536e798e380e4cb1",
                "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
                "raw_sha256": "3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f",
                "adapted_sha256": "6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed",
                "replacements": [
                    {
                        "needle": "\n#[cfg(feature = \"current-api\")]\n",
                        "replacement": "\n#[cfg(feature = \"current-owner-api\")]\n",
                        "count": 2,
                    },
                    {
                        "needle": "    #[cfg(feature = \"current-api\")]\n",
                        "replacement": "    #[cfg(feature = \"current-owner-api\")]\n",
                        "count": 3,
                    },
                    {
                        "needle": "    #[cfg(feature = \"current-api\")]\n",
                        "replacement": "    #[cfg(feature = \"current-property-api\")]\n",
                        "count": 2,
                    },
                    {
                        "needle": "\n#[cfg(all(feature = \"legacy-owner-api\", not(feature = \"current-api\")))]\n",
                        "replacement": "\n#[cfg(all(feature = \"legacy-owner-api\", not(feature = \"current-owner-api\")))]\n",
                        "count": 1,
                    },
                    {
                        "needle": "    #[cfg(all(feature = \"legacy-owner-api\", not(feature = \"current-api\")))]\n",
                        "replacement": "    #[cfg(all(feature = \"legacy-owner-api\", not(feature = \"current-owner-api\")))]\n",
                        "count": 3,
                    },
                    {
                        "needle": "    #[cfg(all(feature = \"legacy-property-api\", not(feature = \"current-api\")))]\n",
                        "replacement": "    #[cfg(all(feature = \"legacy-property-api\", not(feature = \"current-property-api\")))]\n",
                        "count": 2,
                    },
                ],
            },
            "candidate": {
                "source_fence": "fdb139197d08677d6fa0e14d536e798e380e4cb1",
                "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
                "needle": "OleFile::open(Cursor::new(output))?",
                "replacement": "OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default())?",
                "before_sha256": "6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed",
                "after_sha256": "e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782",
            },
        },
        "fixture_relatives": (),
        "fixture_id": "ole-w5-generated-corpus-v1",
        "output_bytes": {"read": 3, "edit": 3},
    },
    "legacy-xls-properties": {
        "probe_relative": "crates/litchi-ops/examples/ole_metadata_bench.rs",
        "lock_relative": "Cargo.lock",
        "probe_sha256": {
            "before": "b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173",
            "after": "53b26136e915fa0bf88518fd1c05e494d55ce5e8f91012be04ac7adf476c1f10",
        },
        "fixture_relatives": (
            "test-data/ole/doc/FancyFoot.doc",
            "test-data/ole/xls/Simple.xls",
            "test-data/ole/ppt/SampleShow.ppt",
        ),
        "fixture_id": "ole-metadata-fixtures-v1",
        "fixture_bytes": {"doc": 27136, "xls": 13824, "ppt": 125440},
        "fixture_sha256": {
            "doc": "709fe8b695a9b12be35c7e81d93beb95afc8dc36422dfc0cbd605a3057d4ddc2",
            "xls": "2b5ec8298bacffdca61c50f6f7607da0a4e63cf76fa011c62f203e6a66bc072a",
            "ppt": "64275a6685b8f267827178037dcc59de927475e56385725916f4c6039c7ce7f6",
        },
        "repeats": 50,
    },
}

RTF_ROUTE_CONTRACT: dict[str, dict[str, str]] = {
    "open_1k": {"outcome": "ok", "fidelity": "exact"},
    "noop_1k": {"outcome": "ok", "fidelity": "exact"},
    "text_commit_1k": {"outcome": "ok", "fidelity": "exact"},
    "text_commit_save_1k": {"outcome": "ok", "fidelity": "exact"},
    "replay_1k": {"outcome": "ok", "fidelity": "exact"},
    "tail_commit_1k": {"outcome": "ok", "fidelity": "exact"},
    "unsupported_content": {
        "outcome": "refused",
        "refusal_kind": "unsupported_content",
        "operation": "edit",
        "reason": "underline edits require non-empty text within one paragraph",
        "fidelity": "unchanged",
    },
    "protected_commit": {
        "outcome": "refused",
        "refusal_kind": "protected",
        "operation": "edit",
        "reason": "document protection refuses changed publication",
        "fidelity": "unchanged",
    },
}

CONTRACTS: dict[str, dict[str, Any]] = {
    "legacy-rtf": {
        "baseline": "86d708878fa537eab42fd4bd5ffbc0f79c606191",
        "blocks": 5,
        "programs": ("routes",),
        "routes": (
            "open_1k",
            "noop_1k",
            "text_commit_1k",
            "text_commit_save_1k",
            "replay_1k",
            "tail_commit_1k",
            "unsupported_content",
            "protected_commit",
        ),
        "probe_contract": "rtf-refusal-contract-v1 workload with native-era error variants normalized to one fixed reason/fidelity contract",
    },
    "legacy-ole-properties": {
        "baseline": "604822b44107999aaddb0bffb96b705e9464f546",
        "blocks": 1,
        "programs": ("read", "edit"),
        "routes": ("read", "edit"),
        "samples": 500,
        "probe_contract": "fdb139 reviewed six-property read/edit corpus with fixture identity and exact section output bytes",
    },
    "legacy-xls-properties": {
        "baseline": "509bf28e7531ad75f278eaec3a732714a243915a",
        "blocks": 1,
        "programs": ("metadata",),
        "routes": ("doc", "xls", "ppt"),
        "probe_contract": "ole_metadata_bench fixed fixture IDs, exact input bytes and native/unified byte-equivalence checks",
    },
}

# The XLS source-parent ref is derived from the lazy-entry record.  Keeping it
# in one place makes an accidental substitution of a later lazy source fail
# closed.  The record remains the authority for the derivation rationale.
class ValidationError(ValueError):
    """A receipt or route contract failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink(), f"symlink is forbidden: {path}")
    _expect(path.is_file(), f"file is missing: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _reject_symlink_components(path: Path, label: str) -> None:
    """Reject the raw path before resolution can hide a symlink."""

    cursor = Path(path.anchor)
    for component in path.parts[1:]:
        cursor /= component
        _expect(not cursor.is_symlink(), f"{label} path contains a symlink: {cursor}")


def _read_json(path: Path) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"JSON file is missing or symlinked: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read JSON {path}: {error}") from error
    _expect(isinstance(value, dict), f"JSON root is not an object: {path}")
    return value


def _read_events(path: Path) -> list[dict[str, Any]]:
    _expect(not path.is_symlink() and path.is_file(), f"event journal is missing or symlinked: {path}")
    records: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid event JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"event at {path}:{line_number} is not an object")
        records.append(value)
    return records


def _validate_hash_record(
    record: Any,
    label: str,
    *,
    require_nonempty: bool = True,
    require_executable: bool = False,
    expected_path: Path | None = None,
    under: Path | None = None,
) -> dict[str, Any]:
    _expect(isinstance(record, dict), f"{label} is not an object")
    raw_path = record.get("path")
    _expect(isinstance(raw_path, str) and raw_path, f"{label} path is missing")
    path = Path(raw_path)
    _expect(path.is_absolute(), f"{label} path must be absolute: {path}")
    _reject_symlink_components(path, label)
    path = path.resolve(strict=False)
    if expected_path is not None:
        _expect(path == expected_path.resolve(), f"{label} path is not the reviewed input: {path}")
    if under is not None:
        _expect(path != under and under.resolve() in path.parents, f"{label} path escapes its bound directory: {path}")
    digest, size = _sha256(path)
    _expect(record.get("sha256") == digest, f"{label} SHA-256 changed: {path}")
    _expect(record.get("bytes") == size, f"{label} byte count changed: {path}")
    _expect(HASH.fullmatch(digest) is not None and (size > 0 or not require_nonempty), f"{label} has an invalid identity")
    if require_executable:
        _expect(os.access(path, os.X_OK), f"{label} is not executable: {path}")
        with path.open("rb") as stream:
            _expect(stream.read(4) == b"\x7fELF", f"{label} is not an ELF executable: {path}")
    return {"path": str(path), "bytes": size, "sha256": digest}


def _identity_file_map(identity: dict[str, Any], label: str) -> dict[str, dict[str, Any]]:
    files = identity.get("identity_files")
    _expect(isinstance(files, list) and files, f"{label} source identity files are missing")
    result: dict[str, dict[str, Any]] = {}
    source_root = Path(str(identity.get("source_root", ""))).resolve()
    for item in files:
        record = _validate_hash_record(item, f"{label} identity file", under=source_root)
        _expect(record["path"] not in result, f"{label} identity contains a duplicate file")
        result[record["path"]] = record
    return result


def _validate_preparation_receipt(
    receipt: dict[str, Any], lane: str, manifest_path: Path, identities: dict[str, Any]
) -> dict[str, Any]:
    manifest_data = _read_json(manifest_path)
    _expect(manifest_data.get("schema") == SCHEMA, "qualification manifest schema mismatch")
    _expect(manifest_data.get("lane") == lane, "qualification manifest lane mismatch")
    preparation = receipt.get("preparation")
    _expect(isinstance(preparation, dict), "build receipt preparation binding is missing")
    normalized_preparation = _validate_hash_record(preparation, "preparation receipt")
    path = Path(normalized_preparation["path"])
    digest, size = normalized_preparation["sha256"], normalized_preparation["bytes"]
    prepared = _read_json(path)
    _expect(prepared.get("schema") == "litchi.legacy-native-preparation.v1", "preparation schema mismatch")
    _expect(prepared.get("lane") == lane, "preparation lane mismatch")
    manifest = prepared.get("manifest")
    _expect(isinstance(manifest, dict) and manifest.get("path") == str(manifest_path), "preparation manifest path mismatch")
    manifest_hash, manifest_bytes = _sha256(manifest_path)
    _expect(manifest.get("sha256") == manifest_hash and manifest.get("bytes", manifest_bytes) == manifest_bytes, "preparation manifest hash changed")
    preparer = prepared.get("preparer")
    _expect(isinstance(preparer, dict) and preparer.get("path") == str(PREPARER_PATH), "preparation preparer path mismatch")
    preparer_hash, _ = _sha256(PREPARER_PATH)
    _expect(preparer.get("sha256") == preparer_hash, "preparation preparer hash changed")
    spec = REVIEWED_INPUTS[lane]
    reviewed = prepared.get("reviewed_inputs")
    _expect(isinstance(reviewed, dict), "preparation reviewed inputs are missing")
    _expect(reviewed.get("probe_relative_path") == spec["probe_relative"], "preparation probe binding differs")
    _expect(reviewed.get("probe_sha256") == spec["probe_sha256"], "preparation probe hash binding differs")
    _expect(reviewed.get("adapter", {}) == spec.get("adapter", {}), "preparation consumer-adapter recipe differs")
    _expect(reviewed.get("lock_relative_path") == spec["lock_relative"], "preparation lock binding differs")
    _expect(tuple(reviewed.get("fixture_relative_paths", ())) == tuple(spec["fixture_relatives"]), "preparation fixture binding differs")
    for leg in ("before", "after"):
        expected_dirty = manifest_data.get("legs", {}).get(leg, {}).get("allowed_dirty_paths")
        _expect(reviewed.get("allowed_dirty_paths") == expected_dirty, f"preparation dirty-path allow-list differs for {leg}")
    for leg in ("before", "after"):
        prepared_identity = prepared.get("identities", {}).get(leg)
        receipt_identity = identities.get(leg)
        _expect(isinstance(prepared_identity, dict) and isinstance(receipt_identity, dict), f"preparation identity missing for {leg}")
        for key in (
            "source_root",
            "head",
            "tree",
            "clean",
            "status_paths",
            "status_sha256",
            "diff_sha256",
            "source_tree_sha256",
            "lock_paths",
            "probe_paths",
            "fixture_paths",
            "identity_sha256",
        ):
            _expect(prepared_identity.get(key) == receipt_identity.get(key), f"preparation/source binding differs for {leg}: {key}")
        prepared_files = {
            str(item.get("path")): (item.get("bytes"), item.get("sha256"))
            for item in prepared_identity.get("identity_files", [])
            if isinstance(item, dict)
        }
        receipt_files = {
            str(item.get("path")): (item.get("bytes"), item.get("sha256"))
            for item in receipt_identity.get("identity_files", [])
            if isinstance(item, dict)
        }
        _expect(prepared_files == receipt_files, f"preparation/source files differ for {leg}")
    return {"path": str(path), "bytes": size, "sha256": digest, "schema": prepared.get("schema")}


def _validate_reviewed_inputs(lane: str, identity: dict[str, Any], leg: str) -> dict[str, Any]:
    spec = REVIEWED_INPUTS[lane]
    root = Path(str(identity["source_root"])).resolve()
    expected_lock = (root / str(spec["lock_relative"])).resolve()
    expected_probe = (root / str(spec["probe_relative"])).resolve()
    lock_paths = [str(Path(str(value)).resolve()) for value in identity.get("lock_paths", [])]
    probe_paths = [str(Path(str(value)).resolve()) for value in identity.get("probe_paths", [])]
    fixture_paths = [str(Path(str(value)).resolve()) for value in identity.get("fixture_paths", [])]
    _expect(lock_paths == [str(expected_lock)], f"{leg} lock path is not the reviewed input")
    _expect(probe_paths == [str(expected_probe)], f"{leg} probe path is not the reviewed input")
    expected_fixtures = [str((root / value).resolve()) for value in spec["fixture_relatives"]]
    _expect(fixture_paths == expected_fixtures, f"{leg} fixture paths are not the reviewed inputs")
    files = _identity_file_map(identity, f"{leg}")
    lock = files.get(str(expected_lock))
    _expect(lock is not None, f"{leg} lock file is not bound into source identity")
    if "lock_sha256" in spec:
        _expect(lock["sha256"] == spec["lock_sha256"][leg], f"{leg} lock provenance differs from the reviewed input")
    probe = files.get(str(expected_probe))
    _expect(probe is not None and probe["sha256"] == spec["probe_sha256"][leg], f"{leg} probe is not the reviewed generator")
    if lane == "legacy-rtf":
        adapter = spec["adapter"]
        manifest_path = root / str(adapter["manifest_relative_path"])
        manifest_hash, _ = _sha256(manifest_path)
        _expect(
            manifest_hash == adapter["manifest_sha256"][leg],
            f"{leg} adapter manifest is not the reviewed consumer overlay",
        )
        live = adapter.get("live_current_probe")
        if leg == "after" and live is not None:
            live_path = root / str(live["relative_path"])
            live_hash, _ = _sha256(live_path)
            _expect(live_hash == live["sha256"], f"{leg} live workspace probe is not reviewed")
            lock_source = Path(str(live["lock_source"]))
            lock_hash, _ = _sha256(lock_source)
            _expect(lock_hash == live["lock_sha256"], f"{leg} live workspace lock is not reviewed")
    fixture_records: dict[str, dict[str, Any]] = {}
    for relative in spec["fixture_relatives"]:
        path = (root / relative).resolve()
        record = files.get(str(path))
        _expect(record is not None, f"{leg} fixture is not bound into source identity: {relative}")
        fixture_key = Path(relative).parent.name.lower()
        expected_size = spec.get("fixture_bytes", {}).get(fixture_key)
        expected_sha = spec.get("fixture_sha256", {}).get(fixture_key)
        if expected_size is not None:
            _expect(record["bytes"] == expected_size, f"{leg} fixture byte identity differs: {relative}")
        if expected_sha is not None:
            _expect(record["sha256"] == expected_sha, f"{leg} fixture hash differs: {relative}")
        fixture_records[relative] = record
    return {"lock": lock, "probe": probe, "fixtures": fixture_records}


def _normalize_cargo(argv: list[str]) -> list[str]:
    values = list(argv)
    separator = values.index("--") if "--" in values else len(values)
    before = values[:separator]
    for flag in ("--offline", "--locked"):
        if flag not in before and not any(item.startswith(flag + "=") for item in before):
            before.append(flag)
    return [*before, *values[separator:]]


def _expected_build_argv(lane: str, leg: str) -> list[str]:
    if lane == "legacy-rtf":
        features = "rtf-legacy-api" if leg == "before" else "rtf-current-api"
        return _normalize_cargo(["cargo", "build", "--release", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--no-default-features", "--features", features, "--example", "refusal_contract_bench"])
    if lane == "legacy-ole-properties":
        features = "current-owner-api,legacy-property-api" if leg == "before" else "current-owner-api,current-property-api"
        return _normalize_cargo(["cargo", "build", "--release", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--no-default-features", "--features", features, "--example", "ole_w5_probe"])
    return _normalize_cargo(["cargo", "build", "--release", "--manifest-path", "Cargo.toml", "-p", "litchi-ops", "--no-default-features", "--features", "doc,xls,ppt", "--example", "ole_metadata_bench"])


def _expected_program_args(lane: str, program: str) -> list[str]:
    if lane == "legacy-ole-properties":
        return ["--route", program, "--warmup", "20", "--samples", "500", "--batch", "1"]
    return []


def _expected_report_paths(output_dir: Path, lane: str, program: str, block: int, leg: str, index: int) -> tuple[Path, Path]:
    if lane == "legacy-rtf":
        stem = f"rtf-routes-block-{block}-{leg}-{index}"
    elif lane == "legacy-ole-properties":
        stem = f"ole-{program}-block-{block}-{leg}-{index}"
    else:
        stem = f"xls-properties-block-{block}-{leg}-{index}"
    return output_dir / f"{stem}.jsonl" if lane == "legacy-xls-properties" or lane == "legacy-rtf" else output_dir / f"{stem}.csv", output_dir / f"{stem}.stderr"


def _validate_build_receipt(
    output_dir: Path,
    lane: str,
    contract: dict[str, Any],
    baseline: str,
    candidate: str,
    before_target: Path,
    after_target: Path,
    manifest_path: Path,
) -> tuple[dict[str, Any], dict[tuple[int, int, str], dict[str, Any]]]:
    receipt_path = output_dir / "build-receipt.json"
    receipt = _read_json(receipt_path)
    _expect(receipt.get("schema") == SCHEMA, "build receipt schema mismatch")
    _expect(receipt.get("status") == "passed" and receipt.get("phase") == "build", "build receipt is not passed")
    _expect(receipt.get("lane") == lane, "build receipt lane mismatch")
    _expect(receipt.get("abba_order") == list(ABBA_ORDER), "build receipt ABBA order mismatch")
    _expect(receipt.get("abba_blocks", 1) == contract["blocks"], "build receipt block count mismatch")
    _expect(
        receipt.get("limits")
        == {
            "target_preflight_bytes": TARGET_PRE_FLIGHT,
            "target_hard_bytes": TARGET_HARD,
            "memory_bytes": MEMORY_HARD,
        },
        "build receipt limits are not the fixed qualification caps",
    )
    _expect(manifest_path.is_absolute(), "qualification manifest path must be absolute")
    _reject_symlink_components(manifest_path, "qualification manifest")
    manifest_path = manifest_path.resolve()
    manifest_data = _read_json(manifest_path)
    _expect(manifest_data.get("schema") == SCHEMA, "qualification manifest schema mismatch")
    _expect(manifest_data.get("lane") == lane, "qualification manifest lane mismatch")
    manifest_record = receipt.get("manifest")
    _expect(isinstance(manifest_record, dict) and manifest_record.get("path") == str(manifest_path), "build receipt manifest path is not source-bound")
    manifest_hash, _ = _sha256(manifest_path)
    _expect(manifest_record.get("sha256") == manifest_hash, "build receipt manifest hash is not source-bound")
    runner_record = receipt.get("runner")
    _expect(isinstance(runner_record, dict) and runner_record.get("path") == str(RUNNER_PATH), "build receipt runner path is not source-bound")
    runner_hash, _ = _sha256(RUNNER_PATH)
    _expect(runner_record.get("sha256") == runner_hash, "build receipt runner hash is not source-bound")
    identities = receipt.get("identities")
    _expect(isinstance(identities, dict), "build receipt identities are missing")
    expected_heads = {"before": baseline, "after": candidate}
    expected_targets = {"before": before_target.resolve(), "after": after_target.resolve()}
    identity_summary: dict[str, Any] = {}
    for leg, expected_head in expected_heads.items():
        identity = identities.get(leg)
        _expect(isinstance(identity, dict), f"build receipt identity is missing for {leg}")
        _expect(identity.get("head") == expected_head, f"{leg} source HEAD is not the reviewed baseline/candidate")
        _expect(Path(str(identity.get("target_dir", ""))).resolve() == expected_targets[leg], f"{leg} target identity mismatch")
        for key in ("tree", "source_tree_sha256", "clean", "status_sha256", "diff_sha256", "identity_sha256", "status_paths"):
            _expect(key in identity, f"{leg} source identity omits {key}")
        reviewed_inputs = _validate_reviewed_inputs(lane, identity, leg)
        normalized_files = list(_identity_file_map(identity, f"{leg}").values())
        lock_paths = [str(Path(str(path)).resolve()) for path in identity["lock_paths"]]
        probe_paths = [str(Path(str(path)).resolve()) for path in identity["probe_paths"]]
        fixture_paths = [str(Path(str(path)).resolve()) for path in identity.get("fixture_paths", [])]
        artifacts = identity.get("artifacts")
        _expect(isinstance(artifacts, list) and artifacts, f"{leg} binary artifacts are missing")
        normalized_artifacts = [_validate_hash_record(item, f"{leg} binary artifact", require_executable=True, under=Path(identity["target_dir"])) for item in artifacts]
        for item in normalized_artifacts:
            path = Path(item["path"])
            _expect(path.parent != path and path.is_file(), f"{leg} artifact is not a file: {path}")
        identity_summary[leg] = {
            "head": expected_head,
            "tree": identity.get("tree"),
            "clean": identity.get("clean"),
            "status_paths": identity.get("status_paths"),
            "status_sha256": identity.get("status_sha256"),
            "diff_sha256": identity.get("diff_sha256"),
            "source_tree_sha256": identity.get("source_tree_sha256"),
            "source_identity_sha256": identity.get("identity_sha256"),
            "lock_paths": [str(path) for path in lock_paths],
            "probe_paths": [str(path) for path in probe_paths],
            "fixture_paths": fixture_paths,
            "reviewed_lock": reviewed_inputs["lock"],
            "reviewed_probe": reviewed_inputs["probe"],
            "reviewed_fixtures": reviewed_inputs["fixtures"],
            "identity_files": normalized_files,
            "artifacts": normalized_artifacts,
        }

    preparation_info = _validate_preparation_receipt(receipt, lane, manifest_path, identities)

    build_journal = receipt.get("event_journal")
    _expect(isinstance(build_journal, dict), "build event journal binding is missing")
    build_path = output_dir / "events-build.jsonl"
    _expect(build_journal.get("path") == str(build_path), "build event journal path is not output-bound")
    build_digest, build_bytes = _sha256(build_path)
    _expect(build_journal.get("sha256") == build_digest and build_journal.get("bytes") == build_bytes, "build event journal hash changed")
    build_events = _read_events(build_path)
    _expect(build_journal.get("count") == len(build_events) and receipt.get("event_count") == len(build_events), "build event count differs from receipt")
    _expect(len(build_events) == 2, "build journal does not contain exactly one build per source leg")
    command_logs: set[Path] = set()
    for event, leg in zip(build_events, ("before", "after")):
        identity = identities[leg]
        event_id = f"build-{leg}-{'rtf_refusal_contract' if lane == 'legacy-rtf' else 'ole_properties_probe' if lane == 'legacy-ole-properties' else 'xls_metadata_bench'}"
        _expect(event.get("schema") == SCHEMA and event.get("event_id") == event_id and event.get("leg") == leg, f"build event identity mismatch for {leg}")
        _expect(event.get("argv") == _expected_build_argv(lane, leg), f"build argv differs from reviewed workload for {leg}")
        _expect(event.get("cwd") == str(Path(identity["source_root"]).resolve()), f"build cwd is not source-bound for {leg}")
        _expect(event.get("source_root") == identity.get("source_root") and event.get("target_dir") == identity.get("target_dir"), f"build paths differ from source identity for {leg}")
        _expect(event.get("output_dir") == str(output_dir.resolve()), f"build output directory differs for {leg}")
        _expect(event.get("source_identity_sha256") == identity.get("identity_sha256"), f"build source identity hash differs for {leg}")
        identity_after = event.get("source_identity_after")
        _expect(isinstance(identity_after, dict), f"build source identity-after is missing for {leg}")
        for key in ("head", "tree", "source_tree_sha256", "clean", "status_paths", "status_sha256", "diff_sha256", "identity_sha256", "lock_paths", "probe_paths", "fixture_paths"):
            _expect(identity_after.get(key) == identity.get(key), f"build source identity-after differs for {leg}: {key}")
        _expect(event.get("returncode") == 0 and event.get("runner_error") is None, f"build event failed for {leg}")
        _expect(not event.get("process_guard_violations") and not event.get("post_exit_descendants") and not event.get("post_exit_offenders"), f"build event was not quiet for {leg}")
        environment = event.get("environment")
        _expect(isinstance(environment, dict), f"build environment is missing for {leg}")
        _expect(environment.get("CARGO_TARGET_DIR") == identity.get("target_dir") and environment.get("CARGO_BUILD_JOBS") == "1" and environment.get("CARGO_INCREMENTAL") == "0" and environment.get("LITCHI_MEM_MAX") == "20G", f"build caps differ for {leg}")
        target = event.get("target")
        _expect(isinstance(target, dict) and target.get("path") == identity.get("target_dir"), f"build target is not source-bound for {leg}")
        _expect(isinstance(target.get("bytes_before"), int) and isinstance(target.get("bytes_after"), int) and 0 <= target["bytes_before"] <= TARGET_HARD and 0 <= target["bytes_after"] <= TARGET_HARD, f"build target caps are invalid for {leg}")
        artifact_records = event.get("outputs")
        _expect(isinstance(artifact_records, list), f"build artifacts are missing from event for {leg}")
        expected_artifacts = identity["artifacts"]
        _expect([item.get("path") for item in artifact_records if isinstance(item, dict)] == [item.get("path") for item in expected_artifacts], f"build artifact paths differ for {leg}")
        for event_artifact, identity_artifact in zip(artifact_records, expected_artifacts):
            _expect(event_artifact.get("exists") is True, f"build artifact was not observed for {leg}")
            normalized = _validate_hash_record(event_artifact, f"{leg} build artifact", require_executable=True, under=Path(identity["target_dir"]))
            _expect(normalized["sha256"] == identity_artifact.get("sha256") and normalized["bytes"] == identity_artifact.get("bytes"), f"build artifact identity differs for {leg}")
        for stream_name in ("stdout", "stderr"):
            stream = event.get(stream_name)
            _expect(isinstance(stream, dict), f"build {stream_name} is missing for {leg}")
            stream_path = Path(str(stream.get("path", ""))).resolve()
            _expect(stream_path.parent == (output_dir / "commands").resolve(), f"build {stream_name} escapes command-log directory for {leg}")
            normalized = _validate_hash_record(stream, f"{leg} build {stream_name}", require_nonempty=False, under=(output_dir / "commands"))
            command_logs.add(Path(normalized["path"]))
    _expect((output_dir / "commands").is_dir(), "build command-log directory is missing")
    _expect(set((output_dir / "commands").iterdir()) == command_logs, "build command-log directory contains unbound files")

    event_path = output_dir / "events-abba.jsonl"
    events = _read_events(event_path)
    expected_events: list[tuple[int, int, str, str]] = []
    for block in range(1, int(contract["blocks"]) + 1):
        for index, leg in enumerate(ABBA_ORDER, start=1):
            for program in contract["programs"]:
                expected_events.append((block, index, leg, program))
    _expect(len(events) == len(expected_events), "ABBA event count differs from the fixed workload")
    event_map: dict[tuple[int, int, str], dict[str, Any]] = {}
    expected_heads = {"before": baseline, "after": candidate}
    for event, (block, index, leg, program) in zip(events, expected_events):
        expected_id = (
            f"abba-{index:02d}-{leg}-{program}"
            if contract["blocks"] == 1
            else f"abba-block-{block:02d}-{index:02d}-{leg}-{program}"
        )
        _expect(event.get("event_id") == expected_id, f"unexpected ABBA event order for {program}")
        _expect(event.get("schema") == SCHEMA and event.get("leg") == leg, f"ABBA event identity mismatch for {program}")
        _expect(event.get("source_head") == expected_heads[leg], f"ABBA source identity mismatch for {program}")
        identity = identities[leg]
        _expect(event.get("source_root") == identity.get("source_root") and event.get("target_dir") == identity.get("target_dir"), f"ABBA source/target binding differs for {program}")
        _expect(event.get("output_dir") == str(output_dir.resolve()), f"ABBA output binding differs for {program}")
        identity_after = event.get("source_identity_after")
        _expect(isinstance(identity_after, dict), f"ABBA source identity-after is missing for {program}")
        for key in ("head", "tree", "source_tree_sha256", "clean", "status_paths", "status_sha256", "diff_sha256", "identity_sha256", "lock_paths", "probe_paths", "fixture_paths"):
            _expect(identity_after.get(key) == identity.get(key), f"ABBA source identity-after differs for {program}: {key}")
        _expect(event.get("returncode") == 0 and event.get("runner_error") is None, f"ABBA child failed for {program}")
        _expect(not event.get("process_guard_violations"), f"ABBA host was not quiet for {program}")
        _expect(not event.get("post_exit_descendants") and not event.get("post_exit_offenders"), f"ABBA descendants/offenders remained for {program}")
        environment = event.get("environment")
        _expect(isinstance(environment, dict), f"ABBA environment missing for {program}")
        _expect(
            environment.get("CARGO_BUILD_JOBS") == "1"
            and environment.get("CARGO_INCREMENTAL") == "0"
            and environment.get("LITCHI_MEM_MAX") == "20G",
            f"ABBA caps differ for {program}",
        )
        _expect("LITCHI_CAPPED" not in environment and "LITCHI_BLESS" not in environment, f"ABBA inherited a qualification override for {program}")
        target = event.get("target")
        _expect(isinstance(target, dict), f"ABBA target receipt missing for {program}")
        _expect(
            Path(str(target.get("path", ""))).resolve()
            == {"before": before_target.resolve(), "after": after_target.resolve()}[leg],
            f"ABBA target identity mismatch for {program}",
        )
        _expect(
            isinstance(target.get("bytes_after"), int)
            and 0 <= target["bytes_after"] <= TARGET_HARD,
            f"ABBA target cap failed for {program}",
        )
        artifacts = identity.get("artifacts", [])
        _expect(len(artifacts) == 1, f"{program} has an ambiguous binary identity")
        binary = Path(str(artifacts[0]["path"])).resolve()
        expected_report, expected_stderr = _expected_report_paths(output_dir, lane, program, block, leg, index)
        _expect(event.get("cwd") == str(Path(identity["source_root"]).resolve()), f"ABBA cwd is not source-bound for {program}")
        _expect(event.get("argv") == [str(binary), *_expected_program_args(lane, program)], f"ABBA argv differs from reviewed workload for {program}")
        _expect(event.get("outputs") == [], f"ABBA declared output files are not empty for {program}")
        stdout_record = event.get("stdout")
        stderr_record = event.get("stderr")
        _expect(isinstance(stdout_record, dict) and Path(str(stdout_record.get("path", ""))).resolve() == expected_report.resolve(), f"{program} stdout path is not workload-bound")
        _expect(isinstance(stderr_record, dict) and Path(str(stderr_record.get("path", ""))).resolve() == expected_stderr.resolve(), f"{program} stderr path is not workload-bound")
        stdout = _validate_hash_record(stdout_record, f"{program} stdout", under=output_dir)
        _validate_hash_record(stderr_record, f"{program} stderr", require_nonempty=False, under=output_dir)
        event_map[(block, index, program)] = {"event": event, "stdout": Path(stdout["path"])}
    return {"receipt": receipt, "identities": identity_summary, "preparation": preparation_info}, event_map


def _validate_rtf_log(path: Path, routes: Iterable[str]) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"RTF report is missing or symlinked: {path}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid RTF JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"RTF row is not an object: {path}:{line_number}")
        _expect(value.get("scenario") in routes, f"unexpected RTF route in {path}: {value.get('scenario')!r}")
        _expect(value.get("workload") == "rtf-refusal-contract-v1", f"RTF workload identity is missing in {path}")
        expected_contract = RTF_ROUTE_CONTRACT[str(value["scenario"])]
        for field, expected_value in expected_contract.items():
            _expect(value.get(field) == expected_value, f"RTF {value['scenario']} {field} differs from the fixed refusal contract")
        median = value.get("median_us")
        _expect(isinstance(median, (int, float)) and math.isfinite(float(median)) and median > 0, f"invalid RTF median in {path}")
        rows.append({"scenario": value["scenario"], "median_us": float(median), "outcome": value["outcome"], "fidelity": value["fidelity"]})
    expected = list(routes)
    _expect([row["scenario"] for row in rows] == expected, f"RTF route/output set differs in {path}")
    return {"path": str(path), "bytes": _sha256(path)[1], "sha256": _sha256(path)[0], "routes": expected, "fixture_id": "rtf-refusal-contract-v1"}


def _validate_csv_log(path: Path, route: str, samples: int, *, expected_output_bytes: int | None = None) -> dict[str, Any]:
    _expect(not path.is_symlink() and path.is_file(), f"CSV report is missing or symlinked: {path}")
    with path.open(newline="", encoding="utf-8") as stream:
        rows = list(csv.reader(stream))
    _expect(rows and rows[0] == ["route", "sample", "nanoseconds", "output_bytes"], f"CSV header mismatch in {path}")
    data = rows[1:]
    _expect(len(data) == samples, f"CSV sample count differs in {path}")
    for expected_sample, row in enumerate(data):
        _expect(len(row) == 4 and row[0] == route, f"CSV route shape differs in {path}")
        try:
            sample, elapsed, output_bytes = int(row[1]), int(row[2]), int(row[3])
        except ValueError as error:
            raise ValidationError(f"CSV numeric field is invalid in {path}") from error
        _expect(sample == expected_sample and elapsed >= 0 and (output_bytes == expected_output_bytes if expected_output_bytes is not None else output_bytes > 0), f"CSV sample/output check failed in {path}")
    digest, size = _sha256(path)
    return {"path": str(path), "bytes": size, "sha256": digest, "route": route, "samples": samples, "fixture_id": "ole-w5-generated-corpus-v1", "expected_output_bytes": expected_output_bytes}


def _validate_metadata_log(path: Path, cases: Iterable[str], *, expected_bytes: dict[str, int] | None = None, expected_repeats: int | None = None) -> dict[str, Any]:
    """Validate the existing metadata benchmark's three deterministic rows."""

    _expect(not path.is_symlink() and path.is_file(), f"metadata report is missing or symlinked: {path}")
    rows: list[dict[str, Any]] = []
    for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
        try:
            value = json.loads(line)
        except json.JSONDecodeError as error:
            raise ValidationError(f"invalid metadata JSON at {path}:{line_number}: {error}") from error
        _expect(isinstance(value, dict), f"metadata row is not an object: {path}:{line_number}")
        _expect(value.get("case") in cases, f"unexpected metadata case in {path}: {value.get('case')!r}")
        _expect(
            value.get("bytes") == expected_bytes.get(str(value["case"])) if expected_bytes is not None else isinstance(value.get("bytes"), int) and value["bytes"] > 0,
            f"metadata fixture/output bytes differ in {path}",
        )
        _expect(
            value.get("repeats") == expected_repeats if expected_repeats is not None else isinstance(value.get("repeats"), int) and value["repeats"] > 0,
            f"metadata workload repeat count differs in {path}",
        )
        for field in ("native_us", "unified_us", "overhead_pct"):
            number = value.get(field)
            _expect(isinstance(number, (int, float)) and math.isfinite(float(number)), f"invalid metadata {field} in {path}")
        _expect(value["native_us"] > 0 and value["unified_us"] > 0, f"metadata elapsed values are not positive in {path}")
        rows.append({"case": value["case"], "bytes": value["bytes"], "repeats": value["repeats"]})
    expected = list(cases)
    _expect([row["case"] for row in rows] == expected, f"metadata route/output set differs in {path}")
    digest, size = _sha256(path)
    return {"path": str(path), "bytes": size, "sha256": digest, "cases": expected, "fixture_id": "ole-metadata-fixtures-v1", "expected_bytes": expected_bytes, "repeats": expected_repeats}


def validate_directory(
    output_dir: Path,
    lane: str,
    baseline: str,
    candidate: str,
    before_target: Path,
    after_target: Path,
    manifest_path: Path,
) -> dict[str, Any]:
    _expect(lane in CONTRACTS, f"unknown legacy native lane: {lane}")
    contract = CONTRACTS[lane]
    _expect(baseline == contract["baseline"], f"{lane} baseline is not the original source fence")
    _expect(FULL_SHA.fullmatch(candidate) is not None and candidate != baseline, "candidate must be a distinct full source SHA")
    receipt_info, events = _validate_build_receipt(
        output_dir, lane, contract, baseline, candidate, before_target, after_target, manifest_path
    )
    reports: list[dict[str, Any]] = []
    for block in range(1, int(contract["blocks"]) + 1):
        for index, leg in enumerate(ABBA_ORDER, start=1):
            for program in contract["programs"]:
                stdout = events[(block, index, program)]["stdout"]
                if lane == "legacy-rtf":
                    report = _validate_rtf_log(stdout, contract["routes"])
                elif lane == "legacy-xls-properties":
                    report = _validate_metadata_log(
                        stdout,
                        contract["routes"],
                        expected_bytes=REVIEWED_INPUTS[lane]["fixture_bytes"],
                        expected_repeats=int(REVIEWED_INPUTS[lane]["repeats"]),
                    )
                else:
                    route = {
                        "read": "read",
                        "edit": "edit",
                    }[program]
                    report = _validate_csv_log(
                        stdout,
                        route,
                        int(contract["samples"]),
                        expected_output_bytes=int(REVIEWED_INPUTS[lane]["output_bytes"][route]),
                    )
                report.update({"block": block, "leg": leg, "program": program})
                reports.append(report)
    return {
        "schema": SUMMARY_SCHEMA,
        "lane": lane,
        "baseline_source": baseline,
        "candidate_source": candidate,
        "abba_order": list(ABBA_ORDER),
        "abba_blocks": contract["blocks"],
        "routes": list(contract["routes"]),
        "probe_contract": contract["probe_contract"],
        "reports": reports,
        "source_lock_probe_binary_identity": receipt_info["identities"],
        "preparation_receipt": receipt_info["preparation"],
        "source_binding": {
            "manifest": str(manifest_path.resolve()),
            "runner": str(RUNNER_PATH),
            "preparer": str(PREPARER_PATH),
            "reviewed_probe_sha256": REVIEWED_INPUTS[lane]["probe_sha256"],
            "consumer_adapter": REVIEWED_INPUTS[lane].get("adapter", {}),
            "fixture_id": REVIEWED_INPUTS[lane]["fixture_id"],
        },
        "quiet_and_caps": {
            "process_guard": True,
            "target_preflight_bytes": TARGET_PRE_FLIGHT,
            "target_hard_bytes": TARGET_HARD,
            "memory_bytes": MEMORY_HARD,
        },
        "allocation": {
            "normal_timing_instrumentation": False,
            "separate_phase_required": True,
            "elapsed_claimable": False,
        },
        "performance_claim": False,
    }


def main(argv: list[str]) -> int:
    if len(argv) != 7:
        print(
            "usage: legacy_native_validator.py OUTPUT_DIR LANE BASELINE_SHA "
            "CANDIDATE_SHA BEFORE_TARGET AFTER_TARGET MANIFEST_PATH",
            file=sys.stderr,
        )
        return 2
    try:
        summary = validate_directory(
            Path(argv[0]),
            argv[1],
            argv[2],
            argv[3],
            Path(argv[4]),
            Path(argv[5]),
            Path(argv[6]),
        )
        destination = Path(argv[0]) / "summary.json"
        _expect(not destination.exists() and not destination.is_symlink(), "summary.json already exists")
        with destination.open("x", encoding="utf-8") as stream:
            json.dump(summary, stream, indent=2, sort_keys=True)
            stream.write("\n")
        print(f"validated {len(summary['reports'])} legacy native reports")
    except (OSError, ValidationError) as error:
        print(f"legacy native qualification validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
