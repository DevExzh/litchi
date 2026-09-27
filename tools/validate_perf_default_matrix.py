"""Validate a default performance report against its checked matrix manifest.

The release harness has one default ``Case`` selection, but it expands that
selection over several independent corpus dimensions.  Keeping the expected
row identities in the checked manifest lets CI validate the complete matrix
without repeating a count that can go stale when a default case is added.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any


class MatrixValidationError(ValueError):
    """The checked matrix manifest or report is malformed or incomplete."""


MANIFEST_KEYS = {
    "schema_version",
    "manifest_kind",
    "harness_source",
    "report_schema_version",
    "source_report_samples_per_case",
    "source_report_warmup_iterations_per_case",
    "result_count",
    "case_count",
    "default_cases",
    "identity_configuration",
    "case_corpora",
    "corpora",
    "canonicalization",
    "result_keys_sha256",
}
IDENTITY_LIST_FIELDS = (
    "corpus_shapes",
    "payload_kinds",
    "writer_shapes",
    "xlsx_shapes",
    "semantic_shapes",
    "rtf_variants",
)
IDENTITY_FIXED_FIELDS = (
    "filesystem_cache_states",
    "filesystem_fresh_child_per_sample",
    "filesystem_process_isolated",
    "filesystem_root_selected",
    "range_simulation",
)
RANGE_SIMULATION_KEYS = {
    "fixed_latency_us",
    "request_overhead_us",
    "bandwidth_bytes_per_second",
    "max_physical_range_bytes",
}
CORPUS_BASE_KEYS = {
    "name",
    "generator",
    "package_format",
    "shape",
    "payload_kind",
    "compression",
    "entry_count",
    "archive_member_count",
    "entry_bytes",
    "uncompressed_payload_bytes",
    "archive_bytes",
    "archive_sha256",
    "target_entry",
    "target_payload_bytes",
    "target_payload_sha256",
    "xlsx",
}
CORPUS_OPTIONAL_KEYS = {"rtf_variant"}
SINK_BUCKET_KEYS = {
    "bytes_0",
    "bytes_1_to_512",
    "bytes_513_to_4096",
    "bytes_4097_to_16384",
    "bytes_16385_to_65536",
    "bytes_over_65536",
}


def _read_json(path: Path, label: str) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError, TypeError) as error:
        raise MatrixValidationError(f"cannot read {label} {path}: {error}") from error


def _object(value: Any, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise MatrixValidationError(f"{context} must be an object")
    return value


def _string(value: Any, context: str) -> str:
    if not isinstance(value, str) or not value:
        raise MatrixValidationError(f"{context} must be a non-empty string")
    return value


def _integer(value: Any, context: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise MatrixValidationError(f"{context} must be an integer >= {minimum}")
    return value


def _boolean(value: Any, context: str) -> bool:
    if not isinstance(value, bool):
        raise MatrixValidationError(f"{context} must be a boolean")
    return value


def _strings(value: Any, context: str) -> list[str]:
    if not isinstance(value, list) or not value:
        raise MatrixValidationError(f"{context} must be a non-empty string list")
    result = [_string(item, f"{context}[{index}]") for index, item in enumerate(value)]
    if len(result) != len(set(result)):
        raise MatrixValidationError(f"{context} must be unique")
    return result


def _canonical_json(value: Any) -> str:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    except (TypeError, ValueError, OverflowError) as error:
        raise MatrixValidationError(f"value is not canonical JSON: {error}") from error


def _result_key_digest(keys: set[tuple[str, str]]) -> str:
    raw = b"".join(
        case.encode("utf-8") + b"\0" + corpus.encode("utf-8") + b"\n"
        for case, corpus in sorted(keys)
    )
    return hashlib.sha256(raw).hexdigest()


def load_manifest(path: Path) -> dict[str, Any]:
    """Load and fail closed on the default case/corpus identity manifest."""

    manifest = _object(_read_json(path, "matrix manifest"), "matrix manifest")
    if set(manifest) != MANIFEST_KEYS:
        raise MatrixValidationError("matrix manifest has unexpected or missing fields")
    if _integer(manifest["schema_version"], "manifest schema_version", minimum=1) != 1:
        raise MatrixValidationError("manifest schema_version must be 1")
    if manifest["manifest_kind"] != "case-corpus-key-identity":
        raise MatrixValidationError("manifest kind is not the default case identity")
    if manifest["harness_source"] != "tools/perf-baseline/src/main.rs:Case::DEFAULT":
        raise MatrixValidationError("manifest harness source is not Case::DEFAULT")
    if _integer(manifest["report_schema_version"], "manifest report_schema_version") != 1:
        raise MatrixValidationError("manifest report_schema_version must be 1")
    _integer(manifest["source_report_samples_per_case"], "manifest source samples")
    _integer(manifest["source_report_warmup_iterations_per_case"], "manifest source warmups")

    cases = _strings(manifest["default_cases"], "manifest default_cases")
    case_count = _integer(manifest["case_count"], "manifest case_count", minimum=1)
    result_count = _integer(manifest["result_count"], "manifest result_count", minimum=1)
    if case_count != len(cases):
        raise MatrixValidationError("manifest case_count does not match default_cases")

    identity = _object(manifest["identity_configuration"], "manifest identity_configuration")
    for field in IDENTITY_LIST_FIELDS:
        _strings(identity.get(field), f"manifest identity_configuration.{field}")
    _strings(
        identity.get("filesystem_cache_states"),
        "manifest identity_configuration.filesystem_cache_states",
    )
    for field in (
        "filesystem_fresh_child_per_sample",
        "filesystem_process_isolated",
        "filesystem_root_selected",
    ):
        _boolean(identity.get(field), f"manifest identity_configuration.{field}")
    range_simulation = _object(
        identity.get("range_simulation"),
        "manifest identity_configuration.range_simulation",
    )
    if set(range_simulation) != RANGE_SIMULATION_KEYS:
        raise MatrixValidationError(
            "manifest identity_configuration.range_simulation has unexpected fields"
        )
    for field in RANGE_SIMULATION_KEYS:
        _integer(
            range_simulation[field],
            f"manifest identity_configuration.range_simulation.{field}",
        )

    case_corpora = _object(manifest["case_corpora"], "manifest case_corpora")
    if set(case_corpora) != set(cases):
        raise MatrixValidationError("manifest case_corpora does not match default_cases")
    corpora = _object(manifest["corpora"], "manifest corpora")
    referenced: set[str] = set()
    for case in cases:
        names = _strings(case_corpora[case], f"manifest case_corpora.{case}")
        for name in names:
            referenced.add(name)
            if name not in corpora:
                raise MatrixValidationError(
                    f"manifest case {case} references missing corpus {name}"
                )
    if set(corpora) != referenced:
        raise MatrixValidationError("manifest contains an unreferenced corpus")
    for name, value in corpora.items():
        corpus = _object(value, f"manifest corpus {name}")
        keys = set(corpus)
        if not (keys - CORPUS_OPTIONAL_KEYS) == CORPUS_BASE_KEYS:
            raise MatrixValidationError(f"manifest corpus {name} has unexpected fields")
        if "rtf_variant" in corpus:
            _string(corpus["rtf_variant"], f"manifest corpus {name}.rtf_variant")
        if _string(corpus["name"], f"manifest corpus {name}.name") != name:
            raise MatrixValidationError(f"manifest corpus {name} has a mismatched name")
        _string(corpus["shape"], f"manifest corpus {name}.shape")
        _string(corpus["payload_kind"], f"manifest corpus {name}.payload_kind")

    _string(manifest["canonicalization"], "manifest canonicalization")
    digest = _string(manifest["result_keys_sha256"], "manifest result_keys_sha256")
    if len(digest) != 64 or any(character not in "0123456789abcdef" for character in digest):
        raise MatrixValidationError("manifest result_keys_sha256 is not lowercase SHA-256")

    full_keys = _keys_for_cases(manifest, cases)
    if len(full_keys) != result_count:
        raise MatrixValidationError(
            f"manifest result_count is {result_count}, but its key matrix has {len(full_keys)} rows"
        )
    if _result_key_digest(full_keys) != digest:
        raise MatrixValidationError("manifest result_keys_sha256 does not match its key matrix")
    return manifest


def _keys_for_cases(manifest: dict[str, Any], cases: list[str]) -> set[tuple[str, str]]:
    corpora = manifest["corpora"]
    keys: set[tuple[str, str]] = set()
    for case in cases:
        for name in manifest["case_corpora"][case]:
            key = (case, _canonical_json(corpora[name]))
            if key in keys:
                raise MatrixValidationError(f"duplicate matrix key for {case}: {name}")
            keys.add(key)
    return keys


def expected_keys(
    manifest: dict[str, Any],
    *,
    mode: str,
    shape: str | None = None,
    payload: str | None = None,
) -> set[tuple[str, str]]:
    """Return the full or tiny smoke key matrix from the checked manifest."""

    cases = manifest["default_cases"]
    if mode == "full":
        if shape is not None or payload is not None:
            raise MatrixValidationError("full mode cannot receive smoke selectors")
        return _keys_for_cases(manifest, cases)
    if mode != "smoke" or shape is None or payload is None:
        raise MatrixValidationError("smoke mode requires shape and payload")
    configured_payloads = set(manifest["identity_configuration"]["payload_kinds"])
    if payload not in configured_payloads:
        raise MatrixValidationError(f"smoke payload {payload!r} is absent from the manifest")

    corpora = manifest["corpora"]
    selected: set[tuple[str, str]] = set()
    for case in cases:
        candidates = [
            corpora[name]
            for name in manifest["case_corpora"][case]
            if corpora[name]["shape"] == shape
        ]
        if not candidates:
            raise MatrixValidationError(f"smoke shape {shape!r} has no corpus for {case}")
        payload_candidates = [
            corpus for corpus in candidates if corpus["payload_kind"] in configured_payloads
        ]
        if payload_candidates:
            candidates = [corpus for corpus in payload_candidates if corpus["payload_kind"] == payload]
        if len(candidates) != 1:
            raise MatrixValidationError(
                f"smoke selection for {case} has {len(candidates)} matching corpora"
            )
        key = (case, _canonical_json(candidates[0]))
        if key in selected:
            raise MatrixValidationError(f"duplicate smoke matrix key for {case}")
        selected.add(key)
    return selected


def _validate_configuration(
    report: dict[str, Any],
    manifest: dict[str, Any],
    *,
    mode: str,
    samples: int,
    shape: str | None,
    payload: str | None,
) -> None:
    configuration = _object(report.get("configuration"), "report configuration")
    if configuration.get("cases") != manifest["default_cases"]:
        raise MatrixValidationError("report configuration cases do not match Case::DEFAULT")
    requested_samples = _integer(samples, "requested samples", minimum=1)
    configured_samples = _integer(
        configuration.get("samples_per_case"),
        "report configuration.samples_per_case",
        minimum=1,
    )
    if configured_samples != requested_samples:
        raise MatrixValidationError(
            "report configuration samples_per_case does not match the requested samples"
        )
    identity = manifest["identity_configuration"]
    for field in IDENTITY_FIXED_FIELDS:
        if configuration.get(field) != identity[field]:
            raise MatrixValidationError(
                f"report configuration {field} does not match the manifest"
            )
    if mode == "full":
        for field in IDENTITY_LIST_FIELDS:
            if configuration.get(field) != identity[field]:
                raise MatrixValidationError(
                    f"report configuration {field} does not match the full manifest"
                )
        return
    expected = {
        "corpus_shapes": [shape],
        "payload_kinds": [payload],
        "writer_shapes": [shape],
        "xlsx_shapes": [shape],
        "semantic_shapes": [shape],
    }
    for field, value in expected.items():
        if configuration.get(field) != value:
            raise MatrixValidationError(
                f"smoke report configuration {field} does not match {value!r}"
            )


def _validate_sink(result: dict[str, Any], context: str) -> None:
    sink = result.get("sink")
    if sink is None:
        return
    sink = _object(sink, f"{context} sink")
    buckets = _object(sink.get("write_size_buckets"), f"{context} sink buckets")
    if set(buckets) != SINK_BUCKET_KEYS:
        raise MatrixValidationError(f"{context} sink buckets have unexpected keys")
    write_calls = _integer(sink.get("write_calls"), f"{context} sink write_calls")
    values = [_integer(value, f"{context} sink bucket") for value in buckets.values()]
    if sum(values) != write_calls:
        raise MatrixValidationError(f"{context} sink buckets do not sum to write_calls")


def validate_report(
    report_path: Path,
    manifest: dict[str, Any],
    *,
    mode: str,
    samples: int,
    shape: str | None = None,
    payload: str | None = None,
) -> dict[str, Any]:
    """Validate report identity, configuration, exact keys, and sample shape."""

    samples = _integer(samples, "requested samples", minimum=1)
    report = _object(_read_json(report_path, "performance report"), "performance report")
    if report.get("schema_version") != manifest["report_schema_version"]:
        raise MatrixValidationError("report schema_version does not match the manifest")
    _validate_configuration(
        report,
        manifest,
        mode=mode,
        samples=samples,
        shape=shape,
        payload=payload,
    )

    expected = expected_keys(manifest, mode=mode, shape=shape, payload=payload)
    results = report.get("results")
    if not isinstance(results, list):
        raise MatrixValidationError("report results must be a list")
    actual: set[tuple[str, str]] = set()
    for index, value in enumerate(results):
        result = _object(value, f"report results[{index}]")
        case = _string(result.get("case"), f"report results[{index}].case")
        corpus = _object(result.get("corpus"), f"report results[{index}].corpus")
        key = (case, _canonical_json(corpus))
        if key in actual:
            raise MatrixValidationError(f"report contains duplicate matrix key for {case}")
        actual.add(key)
        elapsed = _object(result.get("elapsed_ns"), f"report results[{index}].elapsed_ns")
        if elapsed.get("unit") != "ns":
            raise MatrixValidationError(
                f"report results[{index}].elapsed_ns.unit must be 'ns'"
            )
        elapsed_samples = elapsed.get("samples")
        if not isinstance(elapsed_samples, list) or len(elapsed_samples) != samples:
            raise MatrixValidationError(
                f"report results[{index}] does not contain exactly {samples} samples"
            )
        for sample_index, sample in enumerate(elapsed_samples):
            _integer(
                sample,
                f"report results[{index}].elapsed_ns.samples[{sample_index}]",
            )
        sample_order = elapsed.get("sample_order")
        if not isinstance(sample_order, list) or len(sample_order) != samples:
            raise MatrixValidationError(
                f"report results[{index}] does not contain exactly {samples} sample_order entries"
            )
        for sample_index, order_index in enumerate(sample_order):
            _integer(
                order_index,
                f"report results[{index}].elapsed_ns.sample_order[{sample_index}]",
            )
        if sorted(sample_order) != list(range(samples)):
            raise MatrixValidationError(
                f"report results[{index}].elapsed_ns.sample_order is not a permutation"
            )
        _validate_sink(result, f"report results[{index}]")
    if actual != expected:
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        detail = []
        if missing:
            detail.append(f"missing {len(missing)} matrix keys")
        if extra:
            detail.append(f"unexpected {len(extra)} matrix keys")
        raise MatrixValidationError("report key matrix mismatch: " + ", ".join(detail))
    if mode == "full":
        if len(actual) != manifest["result_count"]:
            raise MatrixValidationError("full report count does not match manifest result_count")
        if _result_key_digest(actual) != manifest["result_keys_sha256"]:
            raise MatrixValidationError("full report key digest does not match manifest")
    return {
        "mode": mode,
        "result_count": len(actual),
        "case_count": len({case for case, _ in actual}),
        "result_keys_sha256": _result_key_digest(actual),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--mode", choices=("full", "smoke"), required=True)
    parser.add_argument("--samples", type=int, required=True)
    parser.add_argument("--shape")
    parser.add_argument("--payload")
    args = parser.parse_args()
    try:
        manifest = load_manifest(args.manifest)
        summary = validate_report(
            args.report,
            manifest,
            mode=args.mode,
            samples=args.samples,
            shape=args.shape,
            payload=args.payload,
        )
    except MatrixValidationError as error:
        parser.exit(1, f"default performance matrix validation failed: {error}\n")
    print(
        "validated default performance matrix: "
        f"{summary['case_count']} cases, {summary['result_count']} rows, "
        f"{summary['result_keys_sha256']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
