#!/usr/bin/env python3
"""Independently validate one completed W5 OLE owner qualification run.

The paired runner is an orchestrator, not the authority that closes a native
qualification.  This validator deliberately re-reads every receipt, journal,
log, executable, oracle artifact, and measurement from disk and repeats the
identity/conservation checks without calling runner parsing helpers.  It is
safe to run after a build or against a build-free fixture; it never invokes
Cargo or a native binary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import sys
from typing import Any, Mapping


SCHEMA = "litchi.ole-owner-qualification.runner.v1"
MANIFEST_SCHEMA = "litchi.ole-owner-qualification.manifest.v1"
MEASUREMENTS_SCHEMA = "litchi.ole-owner-qualification.measurements.v1"
ROLES = ("before", "after")
VARIANTS = ("normal", "allocator")
BUILD_KEYS = tuple(f"{role}/{variant}" for role in ROLES for variant in VARIANTS)
ABBA_ORDER = ("before", "after", "after", "before")
ROUTES = (
    "open", "begin_edit", "noop_commit", "replace_stream", "read", "edit",
    "doc_tracked_revision", "xls_cells", "xls_cells_stage", "doc_history", "xls_history",
)
ALLOCATOR_ROUTES = ("read", "edit", "xls_cells", "xls_cells_stage", "doc_history", "xls_history")
ALLOCATOR_EVIDENCE_FIELDS = (
    "route", "alloc_calls", "alloc_bytes", "dealloc_calls", "dealloc_bytes",
    "realloc_calls", "realloc_old_bytes", "realloc_new_bytes", "base_live_calls",
    "base_live_bytes", "live_calls", "live_bytes", "peak_calls", "peak_bytes",
    "retained_packages", "retained_patches",
)
ORACLE_FILES = (
    "ole-source.bin", "ole-noop.bin", "ole-replacement.bin", "properties.tsv",
    "doc-source.bin", "doc-edit.bin", "xls-source.bin", "xls-edit.bin",
    "oracle.tsv", "allocator.tsv",
)
MIN_SPAN_NS = 250_000_000
BLOCKS = 5
WARMUP = 2
SAMPLES = 3
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_BYTES = 20_000_000_000
MEMORY_BYTES = 20_000_000_000
EXPECTED_FEATURES = {
    "before/normal": "legacy-owner-api,legacy-property-api",
    "before/allocator": "legacy-owner-api,legacy-property-api,w5-counting-allocator",
    "after/normal": "current-api",
    "after/allocator": "current-api,w5-counting-allocator",
}
CLAIM_FREE = {
    "build": "orchestration and source-bound evidence only; no performance acceptance claim",
    "abba": "raw paired evidence only; no performance acceptance claim",
    "measurements": "raw paired evidence only; no performance acceptance claim",
    "oracle": "correctness and identity preflight only; no performance acceptance claim",
}


class ValidationError(RuntimeError):
    """A receipt or measurement is not admissible."""


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise ValidationError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _json(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    digest, size = _sha256(path)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{label} must contain an object")
    return value, digest, size


def _mapping(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise ValidationError(f"{label} must be an object")
    return value


def _same(value: Any, expected: Any, label: str) -> None:
    if value != expected:
        raise ValidationError(f"{label} differs")


def _path(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        raise ValidationError(f"{label} path is missing")
    return Path(value).resolve(strict=False)


def _hash_record(record: Mapping[str, Any], label: str) -> None:
    path = _path(record.get("path"), label)
    digest, size = _sha256(path)
    _same(record.get("sha256"), digest, f"{label} SHA-256")
    _same(record.get("bytes"), size, f"{label} byte count")


def _events(path: Path, label: str) -> tuple[list[dict[str, Any]], str, int]:
    digest, size = _sha256(path)
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
        values = [json.loads(line) for line in lines]
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read {label}: {error}") from error
    if not all(isinstance(value, dict) for value in values):
        raise ValidationError(f"{label} contains a non-object event")
    return values, digest, size


def _event_hash(events: list[Mapping[str, Any]]) -> str:
    return hashlib.sha256(json.dumps(events, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _validate_log(event: Mapping[str, Any], field: str, output_root: Path) -> None:
    record = _mapping(event.get(field), f"event.{field}")
    path = _path(record.get("path"), f"event.{field}")
    try:
        path.relative_to(output_root / "commands")
    except ValueError:
        try:
            path.relative_to(output_root / "runs")
        except ValueError as error:
            raise ValidationError(f"event {field} escapes the output log roots: {path}") from error
    _hash_record(record, f"event.{field}")


def _validate_preparation(preparation: Mapping[str, Any]) -> None:
    _same(preparation.get("schema"), "litchi.ole-owner-qualification.preparation.v1", "preparation schema")
    _same(preparation.get("lane"), "w5-ole-owner", "preparation lane")
    _same(preparation.get("claim_scope"), "external source and consumer preparation only; no Cargo, timing, or acceptance claim", "preparation claim scope")
    overlay = _mapping(preparation.get("baseline_overlay"), "preparation.baseline_overlay")
    _same(overlay.get("allowed_paths"), [], "baseline overlay")
    _same(overlay.get("production_tree"), "unchanged; driver files are external", "baseline production-tree policy")
    shared = _mapping(preparation.get("shared_inputs"), "preparation.shared_inputs")
    _same(shared.get("routes"), list(ROUTES), "preparation routes")
    _same(shared.get("history_retention"), {"doc_steps": 128, "xls_steps": 128, "packages": 129, "patches": 128}, "preparation history retention")
    _same(shared.get("clone_sizes"), {"small_siblings": 1, "large_siblings": 1024}, "preparation clone sizes")
    api = _mapping(shared.get("api_binding"), "preparation API binding")
    _same(api.get("policy"), "the baseline one-argument open and candidate explicit-limits open are checked against source text; no implementation overlay", "preparation API binding policy")
    for endpoint in ROLES:
        source_key = "baseline" if endpoint == "before" else "candidate"
        _mapping(api.get(source_key), f"preparation API binding.{endpoint}")
        source = _mapping(preparation.get(source_key), f"preparation.{endpoint}")
        _same(source.get("clean"), True, f"preparation.{endpoint} clean")
        driver = _mapping(source.get("driver"), f"preparation.{endpoint}.driver")
        files = driver.get("files")
        if not isinstance(files, list) or len(files) != 5:
            raise ValidationError(f"preparation.{endpoint} does not bind five driver inputs")
        names = {Path(str(_mapping(item, "driver file").get("path"))).name for item in files}
        _same(names, {"Cargo.toml", "Cargo.lock", "ole_w5_probe.rs", "ole_package_clone_probe.rs", "ole_owner_oracle.rs"}, f"preparation.{endpoint} driver files")
        for index, item in enumerate(files):
            _hash_record(_mapping(item, "driver file"), f"preparation.{endpoint}.driver[{index}]")
    for name, provenance in (
        ("reviewed_probe", "external matched-harness artifact from 0359j; historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 tools/perf-baseline/examples/ole_w5_probe.rs"),
        ("clone_probe", "external matched-harness artifact from 0359j; historical source a84650cd70390641f7734dc1f3362bcb9cd1e025 tools/perf-baseline/examples/ole_package_clone_probe.rs"),
        ("lock", "external matched-harness lock retained by 0359j W5 qualification; Cargo.lock from the reviewed current driver"),
    ):
        item = _mapping(shared.get(name), f"preparation.{name}")
        _same(item.get("provenance"), provenance, f"preparation.{name} provenance")
        _hash_record(item, f"preparation.{name}")
    oracle = _mapping(shared.get("oracle_probe"), "preparation.oracle_probe")
    _same(oracle.get("provenance"), "tracked source-only correctness oracle; compiled in both endpoint drivers", "preparation.oracle provenance")
    _hash_record(oracle, "preparation.oracle")


def _validate_build(
    manifest: Mapping[str, Any], preparation: Mapping[str, Any], build: Mapping[str, Any],
    build_path: Path, runner_path: Path, manifest_hash: str, preparation_hash: str,
) -> tuple[str, Path, list[dict[str, Any]]]:
    _same(build.get("schema"), SCHEMA, "build receipt schema")
    _same(build.get("phase"), "build", "build receipt phase")
    _same(build.get("status"), "passed", "build receipt status")
    _same(build.get("claim_scope"), CLAIM_FREE["build"], "build claim scope")
    _same(_mapping(build.get("manifest"), "build manifest").get("sha256"), manifest_hash, "build manifest identity")
    _same(_mapping(build.get("preparation"), "build preparation").get("sha256"), preparation_hash, "build preparation identity")
    receipt_runner = _mapping(build.get("runner"), "build runner")
    _same(_path(receipt_runner.get("path"), "build runner"), runner_path, "build runner path")
    _hash_record(receipt_runner, "build runner")
    limits = _mapping(build.get("limits"), "build limits")
    _same(limits, {"target_preflight_bytes": TARGET_PREFLIGHT_BYTES, "target_hard_bytes": TARGET_HARD_BYTES, "memory_bytes": MEMORY_BYTES}, "build limits")
    toolchain = _mapping(build.get("toolchain"), "build toolchain")
    for name in ("cargo", "rustc"):
        record = _mapping(toolchain.get(name), f"toolchain.{name}")
        _hash_record(record, f"toolchain.{name}")
        if not isinstance(record.get("version"), str) or not record["version"].strip():
            raise ValidationError(f"toolchain.{name} version is missing")
    root = build_path.parent
    journal = _mapping(build.get("event_journal"), "build event journal")
    journal_path = _path(journal.get("path"), "build event journal")
    _same(journal_path, root / "events-build.jsonl", "build event journal path")
    events, digest, size = _events(journal_path, "build event journal")
    _same(journal.get("count"), len(events), "build journal count")
    _same(journal.get("sha256"), _event_hash(events), "build journal semantic hash")
    _same(journal.get("file_sha256"), digest, "build journal file hash")
    _same(journal.get("bytes"), size, "build journal byte count")
    _same([event.get("event_id") for event in events], [f"build-{key.replace('/', '-')}" for key in BUILD_KEYS], "build event order")
    builds = _mapping(build.get("builds"), "builds")
    identities = _mapping(build.get("identities"), "build identities")
    for key in BUILD_KEYS:
        role, variant = key.split("/")
        item = _mapping(builds.get(key), f"builds.{key}")
        _same(item.get("features"), EXPECTED_FEATURES[key], f"builds.{key} features")
        event = events[BUILD_KEYS.index(key)]
        _same(event.get("phase"), "build", f"build event {key} phase")
        _same(event.get("role"), role, f"build event {key} role")
        _same(event.get("variant"), variant, f"build event {key} variant")
        identity = _mapping(identities.get(role), f"build identities.{role}")
        _same(event.get("source_identity_sha256"), identity.get("identity_sha256"), f"build event {key} source identity hash")
        _same(event.get("source_identity_before"), identity, f"build event {key} source identity before")
        _same(event.get("source_identity_after"), identity, f"build event {key} source identity after")
        _same(event.get("returncode"), 0, f"build event {key} returncode")
        _same(event.get("runner_error"), None, f"build event {key} runner error")
        if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
            raise ValidationError(f"build event {key} was not quiet")
        expected_environment = {
            "CARGO_TARGET_DIR": _mapping(item, f"builds.{key}").get("target"),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "LITCHI_MEM_MAX": "20G",
        }
        environment = _mapping(event.get("environment"), f"build event {key} environment")
        _same(environment.get("CARGO_TARGET_DIR"), expected_environment["CARGO_TARGET_DIR"], f"build event {key} target environment")
        _same(environment.get("CARGO_BUILD_JOBS"), "1", f"build event {key} Cargo jobs")
        _same(environment.get("CARGO_INCREMENTAL"), "0", f"build event {key} incremental policy")
        _validate_log(event, "stdout", root)
        _validate_log(event, "stderr", root)
        outputs = event.get("outputs")
        artifacts = _mapping(item.get("artifacts"), f"builds.{key}.artifacts")
        if not isinstance(outputs, list) or set(Path(str(value.get("path"))).name for value in outputs if isinstance(value, Mapping)) != {"ole_w5_probe", "ole_package_clone_probe", "ole_owner_oracle"}:
            raise ValidationError(f"build event {key} output set differs")
        for name in ("probe", "clone", "oracle"):
            artifact = _mapping(artifacts.get(name), f"builds.{key}.{name}")
            path = _path(artifact.get("path"), f"builds.{key}.{name}")
            digest, bytes_ = _sha256(path)
            _same(artifact.get("sha256"), digest, f"artifact {key}/{name} hash")
            _same(artifact.get("bytes"), bytes_, f"artifact {key}/{name} size")
            if not path.stat().st_mode & 0o111:
                raise ValidationError(f"artifact {key}/{name} is not executable")
            event_output = next((value for value in outputs if isinstance(value, Mapping) and value.get("path") == str(path)), None)
            if event_output is None or event_output.get("sha256") != digest or event_output.get("bytes") != bytes_ or event_output.get("exists") is not True:
                raise ValidationError(f"build event {key}/{name} artifact handoff differs")
    return _sha256_file(build_path)[0], root, events


def _sha256_file(path: Path) -> tuple[str, int]:
    return _sha256(path)


def _hash_json(value: Mapping[str, Any]) -> str:
    # Receipts bind the bytes on disk, not a re-serialized object.  Callers
    # replace this helper with the path hash below when they have a path.
    raw = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def _validate_measurements(measurements: Mapping[str, Any], measurements_path: Path) -> None:
    _same(measurements.get("schema"), MEASUREMENTS_SCHEMA, "measurements schema")
    _same(measurements.get("status"), "complete; coordinator review pending", "measurements status")
    _same(measurements.get("claim_scope"), CLAIM_FREE["measurements"], "measurements claim scope")
    _same(measurements.get("routes"), list(ROUTES), "measurements routes")
    _same(measurements.get("abba_order"), list(ABBA_ORDER), "measurements ABBA order")
    _same(measurements.get("blocks"), BLOCKS, "measurements block count")
    _same(measurements.get("warmup"), WARMUP, "measurements warmup")
    _same(measurements.get("samples"), SAMPLES, "measurements samples")
    _same(measurements.get("minimum_span_ns"), MIN_SPAN_NS, "measurements minimum span")
    timing = _mapping(measurements.get("timing"), "measurements timing")
    clone_blocks = timing.get("clone")
    if not isinstance(clone_blocks, list) or len(clone_blocks) != BLOCKS:
        raise ValidationError("timing clone does not contain five blocks")
    for block in clone_blocks:
        legs = block.get("legs") if isinstance(block, Mapping) else None
        if not isinstance(legs, list) or [leg.get("role") for leg in legs if isinstance(leg, Mapping)] != list(ABBA_ORDER):
            raise ValidationError("timing clone ABBA legs differ")
        for leg in legs:
            rows = leg.get("rows")
            if not isinstance(rows, list) or len(rows) != 2 * SAMPLES:
                raise ValidationError("timing clone sample count differs")
            if any(int(row.get("nanoseconds", 0)) < MIN_SPAN_NS for row in rows if isinstance(row, Mapping)):
                raise ValidationError("timing clone leg is shorter than minimum span")
    route_reports = timing.get("routes")
    if not isinstance(route_reports, Mapping) or set(route_reports) != set(ROUTES):
        raise ValidationError("timing route set differs")
    for route in ROUTES:
        blocks = route_reports[route]
        if not isinstance(blocks, list) or len(blocks) != BLOCKS:
            raise ValidationError(f"timing route {route} does not contain five blocks")
        for block in blocks:
            legs = block.get("legs") if isinstance(block, Mapping) else None
            if not isinstance(legs, list) or [leg.get("role") for leg in legs if isinstance(leg, Mapping)] != list(ABBA_ORDER):
                raise ValidationError(f"timing route {route} ABBA legs differ")
            for leg in legs:
                rows = leg.get("rows")
                if not isinstance(rows, list) or len(rows) != SAMPLES:
                    raise ValidationError(f"timing route {route} sample count differs")
                if any(int(row.get("nanoseconds", 0)) < MIN_SPAN_NS for row in rows if isinstance(row, Mapping)):
                    raise ValidationError(f"timing route {route} leg is shorter than minimum span")
    allocation = _mapping(measurements.get("allocation"), "measurements allocation")
    _same(allocation.get("abba_order"), list(ABBA_ORDER), "allocation ABBA order")
    _same(allocation.get("elapsed_claimable"), False, "allocation claim policy")
    routes = _mapping(allocation.get("routes"), "allocation routes")
    if set(routes) != set(ALLOCATOR_ROUTES):
        raise ValidationError("allocation route set differs")
    for route in ALLOCATOR_ROUTES:
        legs = _mapping(routes[route], f"allocation.{route}").get("legs")
        if not isinstance(legs, list) or [leg.get("role") for leg in legs if isinstance(leg, Mapping)] != list(ABBA_ORDER):
            raise ValidationError(f"allocation {route} ABBA legs differ")
    _validate_oracle(measurements, measurements_path.parent)


def _validate_oracle(measurements: Mapping[str, Any], output_root: Path) -> None:
    oracle = _mapping(measurements.get("oracle"), "measurements oracle")
    _same(oracle.get("status"), "preflight-validated", "oracle status")
    _same(oracle.get("claim_scope"), CLAIM_FREE["oracle"], "oracle claim scope")
    reports = _mapping(oracle.get("files"), "oracle files")
    if set(reports) != set(BUILD_KEYS):
        raise ValidationError("oracle endpoint set differs")
    selected: dict[str, Mapping[str, Any]] = {}
    for key in BUILD_KEYS:
        report = _mapping(reports[key], f"oracle.{key}")
        directory = _path(report.get("directory"), f"oracle.{key}.directory")
        files = _mapping(report.get("files"), f"oracle.{key}.files")
        if set(files) != set(ORACLE_FILES):
            raise ValidationError(f"oracle.{key} file set differs")
        for name in ORACLE_FILES:
            record = _mapping(files[name], f"oracle.{key}.{name}")
            _hash_record(record, f"oracle.{key}.{name}")
            if _path(record.get("path"), f"oracle.{key}.{name}").parent != directory:
                raise ValidationError(f"oracle.{key}.{name} escapes its endpoint directory")
        metadata = (_path(files["oracle.tsv"]["path"], "oracle metadata")).read_text(encoding="utf-8").splitlines()
        _same(metadata, [
            "schema\t1", "owner_streams\t65", "owner_target\tole_w5_edit_target.bin",
            "owner_original_bytes\t15", "owner_replacement_bytes\t19", "properties\tvalidated",
            "doc\trevision\tInsertion\t0\t24\tW5Oracle", "xls\tCells\t0\t0\t4242",
        ], f"oracle.{key} explicit metadata")
        properties = (_path(files["properties.tsv"]["path"], "oracle properties")).read_text(encoding="utf-8").splitlines()
        _same(properties, [
            "link_base\thttps://example.invalid/base/",
            "hyperlinks\t1",
            "hyperlink\t7\t11\t0\thttps://example.invalid/target\tbookmark",
            "opaque_property\t42\t128",
        ], f"oracle.{key} explicit property values")
        selected[key] = files
        allocation = (_path(files["allocator.tsv"]["path"], "oracle allocator")).read_text(encoding="utf-8").splitlines()
        if key.endswith("/normal"):
            _same(allocation, ["status\tnative-uninstrumented"], f"oracle.{key} native allocator marker")
        else:
            _validate_allocator_rows(allocation, f"oracle.{key}.allocator")
    for key in ("before/allocator", "after/normal", "after/allocator"):
        for name in ORACLE_FILES[:-1]:
            _same(selected[key][name].get("sha256"), selected["before/normal"][name].get("sha256"), f"oracle cross-endpoint digest {key}/{name}")
    for key in BUILD_KEYS:
        _same(selected[key]["ole-source.bin"].get("sha256"), selected[key]["ole-noop.bin"].get("sha256"), f"oracle no-op digest {key}")


def _validate_allocator_rows(lines: list[str], label: str) -> None:
    header = lines[0].split("\t") if lines else []
    expected = [
        "route", "alloc_calls", "alloc_bytes", "dealloc_calls", "dealloc_bytes",
        "realloc_calls", "realloc_old_bytes", "realloc_new_bytes", "base_live_calls",
        "base_live_bytes", "live_calls", "live_bytes", "peak_calls", "peak_bytes",
        "retained_packages", "retained_patches",
    ]
    _same(header, expected, f"{label} header")
    if len(lines) != 1 + len(ALLOCATOR_ROUTES):
        raise ValidationError(f"{label} route count differs")
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != len(expected) or fields[0] not in ALLOCATOR_ROUTES:
            raise ValidationError(f"{label} route row differs")
        values = {key: int(value) for key, value in zip(expected[1:], fields[1:])}
        if any(value < 0 for value in values.values()):
            raise ValidationError(f"{label} has a negative counter")
        expected_calls = values["base_live_calls"] + values["alloc_calls"] - values["dealloc_calls"]
        expected_bytes = (
            values["base_live_bytes"] + values["alloc_bytes"] - values["dealloc_bytes"]
            - values["realloc_old_bytes"] + values["realloc_new_bytes"]
        )
        _same(values["live_calls"], expected_calls, f"{label} live-call conservation")
        _same(values["live_bytes"], expected_bytes, f"{label} live-byte conservation")
        if values["peak_calls"] < values["live_calls"] or values["peak_bytes"] < values["live_bytes"]:
            raise ValidationError(f"{label} peak is below live state")
        expected_retained = (129, 128) if fields[0] in {"doc_history", "xls_history"} else (0, 0)
        _same((values["retained_packages"], values["retained_patches"]), expected_retained, f"{label} retained history count")


def validate(
    *, manifest_path: Path, preparation_path: Path, build_path: Path,
    abba_path: Path, measurements_path: Path, runner_path: Path,
) -> dict[str, Any]:
    manifest, manifest_hash, _ = _json(manifest_path, "manifest")
    _same(manifest.get("schema"), MANIFEST_SCHEMA, "manifest schema")
    _same(manifest.get("lane"), "w5-ole-owner", "manifest lane")
    manifest_allocator = _mapping(manifest.get("allocator"), "manifest allocator")
    _same(manifest_allocator.get("oracle_evidence_file"), "allocator.tsv", "manifest allocator evidence file")
    _same(manifest_allocator.get("oracle_evidence_fields"), list(ALLOCATOR_EVIDENCE_FIELDS), "manifest allocator evidence fields")
    _same(
        manifest_allocator.get("conservation"),
        "live_calls=base_live_calls+alloc_calls-dealloc_calls; live_bytes=base_live_bytes+alloc_bytes-dealloc_bytes-realloc_old_bytes+realloc_new_bytes; peak is at least live; retained history is 129 packages and 128 patches only on history routes",
        "manifest allocator conservation",
    )
    preparation, preparation_hash, _ = _json(preparation_path, "preparation")
    _validate_preparation(preparation)
    build, _, _ = _json(build_path, "build receipt")
    build_file_hash, build_root, _ = _validate_build(
        manifest, preparation, build, build_path, runner_path.resolve(), manifest_hash, preparation_hash
    )
    abba, _, _ = _json(abba_path, "ABBA receipt")
    _same(abba.get("schema"), SCHEMA, "ABBA receipt schema")
    _same(abba.get("phase"), "abba", "ABBA receipt phase")
    _same(abba.get("status"), "passed", "ABBA receipt status")
    _same(abba.get("claim_scope"), "orchestration and source-bound evidence only; no performance acceptance claim", "ABBA claim scope")
    _same(_mapping(abba.get("manifest"), "ABBA manifest").get("sha256"), manifest_hash, "ABBA manifest identity")
    _same(_mapping(abba.get("preparation"), "ABBA preparation").get("sha256"), preparation_hash, "ABBA preparation identity")
    _same(abba.get("build_receipt_sha256"), build_file_hash, "immutable build receipt hash")
    receipt_runner = _mapping(abba.get("runner"), "ABBA runner")
    _hash_record(receipt_runner, "ABBA runner")
    _same(_path(receipt_runner.get("path"), "ABBA runner"), runner_path.resolve(), "ABBA runner path")
    journal = _mapping(abba.get("event_journal"), "ABBA event journal")
    journal_path = _path(journal.get("path"), "ABBA event journal")
    _same(journal_path, build_root / "events-abba.jsonl", "ABBA event journal path")
    events, digest, size = _events(journal_path, "ABBA event journal")
    _same(journal.get("count"), len(events), "ABBA journal count")
    _same(journal.get("sha256"), _event_hash(events), "ABBA journal semantic hash")
    _same(journal.get("file_sha256"), digest, "ABBA journal file hash")
    _same(journal.get("bytes"), size, "ABBA journal byte count")
    for event in events:
        _same(event.get("phase"), "abba", "ABBA event phase")
        _same(event.get("returncode"), 0, "ABBA event returncode")
        _same(event.get("runner_error"), None, "ABBA event runner error")
        if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
            raise ValidationError("ABBA event was not quiet")
        _validate_log(event, "stdout", build_root)
        _validate_log(event, "stderr", build_root)
    measurements, measurements_hash, measurements_bytes = _json(measurements_path, "measurements")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("path"), str(measurements_path.resolve()), "ABBA measurements path")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("sha256"), measurements_hash, "ABBA measurements hash")
    _same(_mapping(abba.get("measurements"), "ABBA measurements").get("bytes"), measurements_bytes, "ABBA measurements byte count")
    _validate_measurements(measurements, measurements_path)
    return {
        "status": "validated",
        "manifest_sha256": manifest_hash,
        "preparation_sha256": preparation_hash,
        "build_receipt_sha256": build_file_hash,
        "measurements_sha256": measurements_hash,
        "abba_events": len(events),
        "claim_scope": "validation only; no performance acceptance claim",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--preparation", type=Path, required=True)
    parser.add_argument("--build-receipt", type=Path, required=True)
    parser.add_argument("--abba-receipt", type=Path, required=True)
    parser.add_argument("--measurements", type=Path, required=True)
    parser.add_argument("--runner", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = validate(
            manifest_path=args.manifest,
            preparation_path=args.preparation,
            build_path=args.build_receipt,
            abba_path=args.abba_receipt,
            measurements_path=args.measurements,
            runner_path=args.runner,
        )
    except (OSError, ValidationError, ValueError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
