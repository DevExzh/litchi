#!/usr/bin/env python3
"""Prepare and attest the source inputs for a legacy native qualification.

This command is intentionally source-only.  It does not invoke Cargo, build a
binary, or run a native probe.  The external qualification setup may apply a
reviewed consumer adapter to a baseline checkout; this script records the
resulting source state and rejects dirty paths outside the lane's explicit
allow-list.  The runner and the receipt validator bind the resulting receipt
to every later build/run event.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any


SCHEMA = "litchi.legacy-native-preparation.v1"
ROOT = Path(__file__).resolve().parents[2]

OLE_SOURCE_FENCE = "fdb139197d08677d6fa0e14d536e798e380e4cb1"
OLE_RAW_SHA256 = "3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f"

# The fdb139 probe has independent owner/property feature guards.  The
# reviewed 604822 baseline has the migrated public Package owner API but the
# legacy property API, so materialization selects those two APIs without
# changing the measured owner/property operations.
OLE_BASELINE_ADAPTER_RECIPE: dict[str, Any] = {
    "source_fence": OLE_SOURCE_FENCE,
    "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
    "raw_sha256": OLE_RAW_SHA256,
    "adapted_sha256": "6890d0524cf4466a96812eb9ba449d4d995b29e1db74888fe42b5a42e8a91aed",
    "replacements": [
        {
            "needle": '\n#[cfg(feature = "current-api")]\n',
            "replacement": '\n#[cfg(feature = "current-owner-api")]\n',
            "count": 2,
        },
        {
            "needle": '    #[cfg(feature = "current-api")]\n',
            "replacement": '    #[cfg(feature = "current-owner-api")]\n',
            "count": 3,
        },
        {
            "needle": '    #[cfg(feature = "current-api")]\n',
            "replacement": '    #[cfg(feature = "current-property-api")]\n',
            "count": 2,
        },
        {
            "needle": '\n#[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]\n',
            "replacement": '\n#[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]\n',
            "count": 1,
        },
        {
            "needle": '    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-api")))]\n',
            "replacement": '    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]\n',
            "count": 3,
        },
        {
            "needle": '    #[cfg(all(feature = "legacy-property-api", not(feature = "current-api")))]\n',
            "replacement": '    #[cfg(all(feature = "legacy-property-api", not(feature = "current-property-api")))]\n',
            "count": 2,
        },
    ],
}

OLE_ADAPTER_RECIPE: dict[str, str] = {
    "source_fence": OLE_SOURCE_FENCE,
    "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
    "needle": "OleFile::open(Cursor::new(output))?",
    "replacement": "OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default())?",
    "before_sha256": OLE_BASELINE_ADAPTER_RECIPE["adapted_sha256"],
    "after_sha256": "e031d0adbc00a5784658c37472693ae1fe840ad88f84a5b785b48dd77a810782",
}

OLE_ADAPTER_RECORD: dict[str, Any] = {
    "source_fence": OLE_SOURCE_FENCE,
    "relative_path": "tools/perf-baseline/examples/ole_w5_probe.rs",
    "baseline": OLE_BASELINE_ADAPTER_RECIPE,
    "candidate": OLE_ADAPTER_RECIPE,
}

RTF_PROBE_SHA256 = "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81"
RTF_BASELINE_LOCK_SHA256 = "57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f"
RTF_CANDIDATE_LOCK_SHA256 = "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c"
RTF_BASELINE_MANIFEST_SHA256 = "af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676"
RTF_CANDIDATE_MANIFEST_SHA256 = "61e103809753ffc528b2cd2c2e00a1525d40fc4e6360668ad982bec22a60bb73"
RTF_LIVE_PROBE_RELATIVE = "crates/litchi-rtf/examples/refusal_contract_bench.rs"
RTF_LIVE_PROBE_SHA256 = "1726acb58953ae3f57d9b74207f75678a3bcc7ac4c993443aac5583d3943f791"

RTF_ADAPTER_RECORD: dict[str, Any] = {
    "relative_path": "tools/perf-baseline/examples/refusal_contract_bench.rs",
    "probe_sha256": RTF_PROBE_SHA256,
    "manifest_relative_path": "tools/perf-baseline/Cargo.toml",
    "manifest_sha256": {
        "before": RTF_BASELINE_MANIFEST_SHA256,
        "after": RTF_CANDIDATE_MANIFEST_SHA256,
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
        "relative_path": RTF_LIVE_PROBE_RELATIVE,
        "sha256": RTF_LIVE_PROBE_SHA256,
        "lock_sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
        "lock_source": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/locks/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock",
        "api": "current litchi-rtf workspace example; not part of the baseline adapter source",
    },
    "historical_failed_compile": {
        "source_sha256": "70077fcec439fda2fdf9e9635f5db2f1f7c86ed31f3f972a95c5e5e3a607a585",
        "source_relative_path": RTF_LIVE_PROBE_RELATIVE,
        "api_context": "current litchi-rtf consumer at candidate fence 35e37dbede6657dd023b114a1536862c18541946",
        "receipt": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/rtf-probe-compile/check-before.json",
        "receipt_sha256": "f3af926353fc3db4f74221f62ed4f4743cc6b31c359ab42415c8f1500e33bda4",
        "log": "/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/rtf-probe-compile/check-before.log",
        "log_sha256": "119d4216675ee79a7e89e17579f019dd4fe3a6d43c08a8bd22d87c37622fd103",
        "exit": 101,
    },
}

LANES: dict[str, dict[str, Any]] = {
    "legacy-rtf": {
        "baseline": "86d708878fa537eab42fd4bd5ffbc0f79c606191",
        "probe": "tools/perf-baseline/examples/refusal_contract_bench.rs",
        "probe_sha256": {
            "before": RTF_PROBE_SHA256,
            "after": RTF_PROBE_SHA256,
        },
        "lock": "tools/perf-baseline/Cargo.lock",
        "fixtures": [],
        "lock_sha256": {
            "before": RTF_BASELINE_LOCK_SHA256,
            "after": RTF_CANDIDATE_LOCK_SHA256,
        },
        "adapter": RTF_ADAPTER_RECORD,
        "allowed_dirty": [
            "tools/perf-baseline/Cargo.lock",
            "tools/perf-baseline/Cargo.toml",
            "tools/perf-baseline/examples/refusal_contract_bench.rs",
        ],
    },
    "legacy-ole-properties": {
        "baseline": "604822b44107999aaddb0bffb96b705e9464f546",
        "probe": "tools/perf-baseline/examples/ole_w5_probe.rs",
        "probe_sha256": {
            "before": OLE_BASELINE_ADAPTER_RECIPE["adapted_sha256"],
            "after": OLE_ADAPTER_RECIPE["after_sha256"],
        },
        "adapter": OLE_ADAPTER_RECORD,
        "lock": "tools/perf-baseline/Cargo.lock",
        "fixtures": [],
        "allowed_dirty": [
            "tools/perf-baseline/Cargo.toml",
            "tools/perf-baseline/Cargo.lock",
            "tools/perf-baseline/examples/ole_w5_probe.rs",
        ],
    },
    "legacy-xls-properties": {
        "baseline": "509bf28e7531ad75f278eaec3a732714a243915a",
        "probe": "crates/litchi-ops/examples/ole_metadata_bench.rs",
        "probe_sha256": {
            "before": "b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173",
            "after": "53b26136e915fa0bf88518fd1c05e494d55ce5e8f91012be04ac7adf476c1f10",
        },
        "lock": "Cargo.lock",
        "fixtures": [
            "test-data/ole/doc/FancyFoot.doc",
            "test-data/ole/xls/Simple.xls",
            "test-data/ole/ppt/SampleShow.ppt",
        ],
        "allowed_dirty": [
            "Cargo.lock",
            "crates/litchi-ops/examples/ole_metadata_bench.rs",
        ],
    },
}


class PreparationError(ValueError):
    """The source checkout does not match the reviewed preparation contract."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _reject_symlink_components(path: Path, label: str) -> None:
    """Reject the raw path before resolution can hide a symlink."""

    raw = Path(path)
    cursor = Path(raw.anchor) if raw.is_absolute() else Path()
    components = raw.parts[1:] if raw.is_absolute() else raw.parts
    for component in components:
        cursor /= component
        _expect(not cursor.is_symlink(), f"{label} path contains a symlink: {cursor}")


def _materialize_ole_file(
    source: Path,
    destination: Path,
    owned_root: Path,
    materializer: Any,
    label: str,
) -> None:
    """Materialize one adapter while keeping its output inside ``owned_root``."""

    source_raw = Path(source)
    destination_raw = Path(destination)
    owned_raw = Path(owned_root)
    _reject_symlink_components(source_raw, f"{label} source")
    _reject_symlink_components(destination_raw, f"{label} destination")
    _reject_symlink_components(owned_raw, f"{label} owned root")
    _expect(source_raw.is_file() and not source_raw.is_symlink(), f"{label} source is missing or symlinked: {source_raw}")
    _expect(owned_raw.is_dir() and not owned_raw.is_symlink(), f"{label} owned root is missing or symlinked: {owned_raw}")
    _expect(not destination_raw.exists() and not destination_raw.is_symlink(), f"refusing to overwrite {label} destination: {destination_raw}")

    source_path = source_raw.resolve(strict=True)
    owned_path = owned_raw.resolve(strict=True)
    destination_path = destination_raw.resolve(strict=False)
    try:
        destination_path.relative_to(owned_path)
    except ValueError as error:
        raise PreparationError(f"{label} destination escapes owned root: {destination_raw}") from error
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    _expect(not destination_path.exists() and not destination_path.is_symlink(), f"refusing to overwrite {label} destination: {destination_path}")
    rendered = materializer(source_path.read_text(encoding="utf-8"))
    try:
        with destination_path.open("x", encoding="utf-8") as stream:
            stream.write(rendered)
    except FileExistsError as error:
        raise PreparationError(f"refusing to overwrite {label} destination: {destination_path}") from error


def materialize_ole_baseline_adapter(source: str) -> str:
    """Select the public-owner/legacy-property APIs for the 604822 baseline."""

    source_bytes = source.encode("utf-8")
    recipe = OLE_BASELINE_ADAPTER_RECIPE
    _expect(hashlib.sha256(source_bytes).hexdigest() == recipe["raw_sha256"], "OLE baseline adapter source is not the fdb139 probe")
    adapted = source
    for replacement in recipe["replacements"]:
        needle = replacement["needle"]
        expected_count = replacement["count"]
        _expect(adapted.count(needle) >= expected_count, "OLE baseline adapter replacement count differs")
        adapted = adapted.replace(needle, replacement["replacement"], expected_count)
    _expect(hashlib.sha256(adapted.encode("utf-8")).hexdigest() == recipe["adapted_sha256"], "OLE baseline adapter output hash differs from the reviewed recipe")
    return adapted


def materialize_ole_adapter(source: str) -> str:
    """Apply the reviewed baseline API split and candidate CFB adapter."""

    baseline = materialize_ole_baseline_adapter(source)
    recipe = OLE_ADAPTER_RECIPE
    _expect(hashlib.sha256(baseline.encode("utf-8")).hexdigest() == recipe["before_sha256"], "OLE candidate adapter input differs from the reviewed baseline adapter")
    _expect(baseline.count(recipe["needle"]) == 1, "OLE adapter needle is not unique")
    adapted = baseline.replace(recipe["needle"], recipe["replacement"])
    _expect(hashlib.sha256(adapted.encode("utf-8")).hexdigest() == recipe["after_sha256"], "OLE adapter output hash differs from the reviewed recipe")
    return adapted


def materialize_ole_adapter_file(source: Path, destination: Path, *, owned_root: Path) -> None:
    """Materialize the candidate adapter under an explicitly owned root."""

    _materialize_ole_file(source, destination, owned_root, materialize_ole_adapter, "OLE adapter")


def materialize_ole_baseline_adapter_file(source: Path, destination: Path, *, owned_root: Path) -> None:
    """Materialize the baseline adapter under an explicitly owned root."""

    _materialize_ole_file(source, destination, owned_root, materialize_ole_baseline_adapter, "OLE baseline adapter")


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _identity_file(root: Path, relative: str) -> dict[str, Any]:
    raw_path = root / relative
    _expect(not raw_path.is_symlink(), f"identity input may not be a symlink: {relative}")
    path = raw_path.resolve()
    _expect(path.parent == (root / relative).parent.resolve(), f"identity path escapes source root: {relative}")
    digest, size = _sha256(path)
    return {"path": str(path), "relative": relative, "bytes": size, "sha256": digest}


def _identity_digest(files: list[dict[str, Any]]) -> str:
    """Use the runner's one-separator-per-field identity framing exactly."""

    digest = hashlib.sha256()
    entries = [
        (Path(record["path"]).name, record)
        for record in files
    ]
    for relative, record in sorted(entries, key=lambda item: (str(Path(item[1]["path"]).parent), item[0])):
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(record["bytes"]).encode("ascii"))
        digest.update(b"\0")
        digest.update(bytes.fromhex(record["sha256"]))
    return digest.hexdigest()


def _source_identity(root: Path, lane: str, leg: str, spec: dict[str, Any]) -> dict[str, Any]:
    root = root.resolve()
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve()
    _expect(top == root, f"source root is not the worktree top-level: {root}")
    head = _git(root, "rev-parse", "HEAD").strip()
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    allowed = set(spec["allowed_dirty"])
    status_paths: list[str] = []
    for line in status.splitlines():
        _expect(len(line) >= 4, f"malformed git status entry in {root}: {line!r}")
        value = line[3:]
        if " -> " in value:
            value = value.split(" -> ", 1)[1]
        value = value.strip()
        status_paths.append(value)
        _expect(value in allowed, f"unreviewed dirty path in {root}: {value}")
    relative_files = [spec["lock"], spec["probe"], *spec["fixtures"]]
    files = [_identity_file(root, relative) for relative in relative_files]
    lock_record = next(record for record in files if record["relative"] == spec["lock"])
    if "lock_sha256" in spec:
        _expect(
            lock_record["sha256"] == spec["lock_sha256"][leg],
            f"{lane}/{leg} lock is not the reviewed provenance lock",
        )
    if lane == "legacy-rtf" and "adapter" in spec:
        adapter = spec["adapter"]
        manifest_path = root / adapter["manifest_relative_path"]
        manifest_sha256, _ = _sha256(manifest_path)
        _expect(
            manifest_sha256 == adapter["manifest_sha256"][leg],
            f"{lane}/{leg} adapter manifest is not the reviewed consumer overlay",
        )
        live = adapter.get("live_current_probe")
        if leg == "after" and live is not None:
            live_path = root / str(live["relative_path"])
            live_sha256, _ = _sha256(live_path)
            _expect(live_sha256 == live["sha256"], f"{lane}/{leg} live workspace probe is not reviewed")
            lock_source = Path(str(live["lock_source"]))
            lock_sha256, _ = _sha256(lock_source)
            _expect(lock_sha256 == live["lock_sha256"], f"{lane}/{leg} live workspace lock is not reviewed")
    probe_record = next(record for record in files if record["relative"] == spec["probe"])
    _expect(probe_record["sha256"] == spec["probe_sha256"][leg], f"{lane}/{leg} probe is not the reviewed generator")
    if lane == "legacy-ole-properties":
        if leg == "before":
            _expect(probe_record["sha256"] == OLE_BASELINE_ADAPTER_RECIPE["adapted_sha256"], "OLE baseline must use the reviewed public-owner/legacy-property adapter")
        else:
            _expect(probe_record["sha256"] == OLE_ADAPTER_RECIPE["after_sha256"], "OLE candidate must use only the reviewed consumer adapter")
    return {
        "source_root": str(root),
        "head": head,
        "tree": tree,
        "clean": not bool(status),
        "status": status,
        "status_paths": status_paths,
        "status_sha256": hashlib.sha256(status.encode("utf-8")).hexdigest(),
        "diff_sha256": hashlib.sha256(_git(root, "diff", "--no-ext-diff", "--binary").encode("utf-8")).hexdigest(),
        "source_tree_sha256": hashlib.sha256(
            (tree + "\0" + status + "\0" + _git(root, "diff", "--no-ext-diff", "--binary")).encode("utf-8")
        ).hexdigest(),
        "lock_paths": [str((root / spec["lock"]).resolve())],
        "probe_paths": [str((root / spec["probe"]).resolve())],
        "fixture_paths": [str((root / value).resolve()) for value in spec["fixtures"]],
        "identity_sha256": _identity_digest(files),
        "identity_files": files,
    }


def prepare(lane: str, manifest: Path, before: Path, after: Path, output: Path) -> None:
    _expect(lane in LANES, f"unknown legacy lane: {lane}")
    spec = LANES[lane]
    _reject_symlink_components(manifest, "qualification manifest")
    _expect(manifest.is_absolute(), "qualification manifest path must be absolute")
    _expect(not manifest.is_symlink(), f"qualification manifest path is symlinked: {manifest}")
    manifest = manifest.resolve()
    manifest_digest, manifest_bytes = _sha256(manifest)
    try:
        manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read qualification manifest: {error}") from error
    _expect(isinstance(manifest_data, dict) and manifest_data.get("lane") == lane, "manifest lane does not match preparation lane")
    _expect(manifest_data.get("schema") == "litchi.native-qualification.v1", "manifest schema is not the runner schema")
    before_identity = _source_identity(before, lane, "before", spec)
    after_identity = _source_identity(after, lane, "after", spec)
    _expect(before_identity["head"] == spec["baseline"], "before source is not the reviewed historical baseline")
    _expect(after_identity["head"] != spec["baseline"], "after source must be a distinct candidate")
    preparer = Path(__file__).resolve()
    preparer_digest, _ = _sha256(preparer)
    receipt = {
        "schema": SCHEMA,
        "lane": lane,
        "manifest": {"path": str(manifest), "sha256": manifest_digest, "bytes": manifest_bytes},
        "preparer": {"path": str(preparer), "sha256": preparer_digest},
        "reviewed_inputs": {
            "probe_relative_path": spec["probe"],
            "probe_sha256": dict(spec["probe_sha256"]),
            "adapter": dict(spec.get("adapter", {})),
            "lock_relative_path": spec["lock"],
            "fixture_relative_paths": list(spec["fixtures"]),
            "allowed_dirty_paths": list(spec["allowed_dirty"]),
        },
        "identities": {"before": before_identity, "after": after_identity},
    }
    output = output.resolve()
    _expect(not output.exists() and not output.is_symlink(), f"refusing to overwrite preparation receipt: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write("\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lane", choices=sorted(LANES), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(args.lane, args.manifest, args.before, args.after, args.output)
    except (OSError, PreparationError) as error:
        print(f"legacy native preparation refused: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
