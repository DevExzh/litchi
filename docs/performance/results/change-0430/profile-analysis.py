#!/usr/bin/env python3
"""Audit retained frame-pointer stacks without making a performance claim.

The perf data in this directory was recorded for CPU attribution.  This tool
only interprets the retained ``perf script`` text and the frozen custody
receipts.  It counts event blocks, rather than weighting blocks by the cycle
number printed in a perf header.  The resulting percentages therefore remain
unweighted full-process sampled-cycle percentages; they are not API timing
fractions.
"""

import argparse
from collections import Counter
from contextlib import contextmanager
import gzip
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile


ROOT = Path(__file__).resolve().parent
SUMMARY_NAME = "profile-summary.json"

CATEGORY_NAMES = ("iteration", "corpus_setup", "other_or_unresolved")
EXPECTED_PROVIDERS = ("bytes", "file")
EXPECTED_RECEIPT_ARTIFACTS = (
    "record.log",
    "report.log",
    "script.log",
    ".data",
    ".json",
)

# perf script's first line for a sample is, for example:
#   litchi-perf-bas 123 214333.092830: 1 cycles:u:
# A count may be much larger than one.  Do not use it as a weight.
SAMPLE_HEADER_RE = re.compile(
    r"^\S+\s+\d+\s+\S+:\s+\d+\s+cycles:u:(?:\s|$)"
)
FRAME_RE = re.compile(
    r"^\s*(?:0x)?[0-9a-fA-F]+\s+(.+?)(?:\s+\([^\n]*\))?\s*$"
)
RECORD_SAMPLES_RE = re.compile(r"\((\d+) samples?\)")
DIAGNOSTIC_RE = re.compile(r"(?:addr2line|could not read first record)", re.I)
ITERATION_RE = re.compile(
    r"(?<![A-Za-z0-9_])run_lifecycle_iteration"
    r"(?:<[^\n>]*>)?(?=\+|$|\s|\()"
)
DEFLATE_MEDIUM_RE = re.compile(
    r"(?<![A-Za-z0-9_])deflate_medium(?:\+|<|$|\s)"
)

SETUP_MARKERS = (
    "build_pptx_source_backed_cross_copy_corpus",
    "build_pptx_cross_copy_corpus",
    "pptx_cross_copy_bytes",
)

API_MARKERS = {
    "plan_cross_slide_copy": "plan_cross_slide_copy",
    "publish_cross_slide_copy_to_stream": "publish_cross_slide_copy_to_stream",
    "from_read_at_with_limits_and_cache_limits_and_execution_context":
        "from_read_at_with_limits_and_cache_limits_and_execution_context",
    "try_cache_diagnostics": "try_cache_diagnostics",
    "sha256_hex": "sha256_hex",
    "rss_point": "rss_point",
}


class AnalysisError(RuntimeError):
    """A retained-bundle or parsing contract failed."""


def fail(message):
    raise AnalysisError(message)


def ensure(condition, message):
    if not condition:
        fail(message)


def reject_duplicate_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            fail("duplicate JSON key: %s" % key)
        result[key] = value
    return result


def sha256_bytes(value):
    return hashlib.sha256(value).hexdigest()


def sha256_path(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def raw_path(root, relative):
    return root / relative


def gzip_path(root, relative):
    return root / (str(relative) + ".gz")


@contextmanager
def open_logical(root, relative):
    """Open the retained raw artifact, whether stored raw or as deterministic gzip."""
    plain = raw_path(root, relative)
    compressed = gzip_path(root, relative)
    ensure(not (plain.is_file() and compressed.is_file()),
           "both raw and gzip forms exist for %s" % relative)
    if plain.is_file():
        with plain.open("rb") as stream:
            yield stream
        return
    if compressed.is_file():
        try:
            with gzip.open(compressed, "rb") as stream:
                yield stream
        except (OSError, EOFError) as error:
            fail("cannot decompress %s: %s" % (compressed, error))
        return
    fail("missing retained artifact %s (or %s)" % (plain, compressed))


def logical_digest(root, relative):
    digest = hashlib.sha256()
    size = 0
    with open_logical(root, relative) as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def read_logical(root, relative):
    with open_logical(root, relative) as stream:
        return stream.read()


def read_raw_json(root, relative):
    path = raw_path(root, relative)
    ensure(path.is_file(), "JSON custody artifact must remain raw: %s" % path)
    try:
        return json.loads(path.read_text(), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail("invalid JSON %s: %s" % (path, error))


def raw_json_bytes(root, relative):
    path = raw_path(root, relative)
    ensure(path.is_file(), "missing raw JSON %s" % path)
    return path.read_bytes()


def verify_logical_artifact(root, relative, expected, compression):
    """Verify receipt's raw hash and, if sealed, the gzip envelope too."""
    expected_sha = expected.get("sha256")
    expected_bytes = expected.get("bytes")
    ensure(isinstance(expected_sha, str) and isinstance(expected_bytes, int),
           "incomplete artifact receipt for %s" % relative)
    actual_sha, actual_bytes = logical_digest(root, relative)
    ensure(actual_sha == expected_sha,
           "raw hash mismatch for %s: %s != %s" %
           (relative, actual_sha, expected_sha))
    ensure(actual_bytes == expected_bytes,
           "raw byte count mismatch for %s: %s != %s" %
           (relative, actual_bytes, expected_bytes))

    plain = raw_path(root, relative)
    compressed = gzip_path(root, relative)
    if compressed.is_file():
        compression_path = root / "compression.json"
        if compression_path.is_file():
            compression_rows = read_raw_json(root, Path("compression.json"))
            row = compression_rows.get(str(relative) + ".gz")
            ensure(isinstance(row, dict),
                   "compression metadata missing for %s" % compressed)
            ensure(row.get("original_path") == str(relative),
                   "compression original path mismatch for %s" % relative)
            ensure(row.get("original_sha256") == expected_sha and
                   row.get("original_bytes") == expected_bytes,
                   "compression raw binding mismatch for %s" % relative)
            ensure(row.get("stored_sha256") == sha256_path(compressed) and
                   row.get("stored_bytes") == compressed.stat().st_size,
                   "compression stored binding mismatch for %s" % compressed)
    else:
        ensure(plain.is_file(), "missing logical artifact %s" % relative)

    return {"path": str(relative), "sha256": expected_sha, "bytes": expected_bytes}


def require_equal(actual, expected, label):
    ensure(actual == expected, "%s mismatch: %r != %r" % (label, actual, expected))


def validate_protocol_and_receipts(root):
    protocol_relative = Path("profile-protocol.json")
    index_relative = Path("profile-index.json")
    protocol_bytes = raw_json_bytes(root, protocol_relative)
    index_bytes = raw_json_bytes(root, index_relative)
    protocol = read_raw_json(root, protocol_relative)
    index = read_raw_json(root, index_relative)
    protocol_sha = sha256_bytes(protocol_bytes)
    index_sha = sha256_bytes(index_bytes)

    require_equal(protocol.get("change"), 430, "protocol change")
    require_equal(protocol.get("event"), "cycles:u", "protocol event")
    require_equal(protocol.get("call_graph"), "fp", "protocol call graph")
    require_equal(protocol.get("corpus"), "media-rich", "protocol corpus")
    require_equal(protocol.get("providers"), list(EXPECTED_PROVIDERS), "protocol providers")
    require_equal(protocol.get("samples"), 100, "protocol samples")
    require_equal(protocol.get("warmup"), 3, "protocol warmup")
    ensure(isinstance(index, list), "profile index is not a list")
    expected_index = ["profiles/%s-receipt.json" % provider
                      for provider in EXPECTED_PROVIDERS]
    require_equal(index, expected_index, "profile index")

    capture_driver_sha = sha256_path(root / "capture-profiles.py")
    verifier_sha = sha256_path(root / "verify-report.py")
    receipts = []
    seen_providers = set()
    for receipt_relative_text in index:
        receipt_relative = Path(receipt_relative_text)
        receipt_bytes = raw_json_bytes(root, receipt_relative)
        receipt = read_raw_json(root, receipt_relative)
        provider = receipt.get("provider")
        ensure(provider in EXPECTED_PROVIDERS and provider not in seen_providers,
               "unexpected or duplicate provider receipt: %s" % provider)
        seen_providers.add(provider)
        require_equal(receipt.get("status"), "pass", "%s receipt status" % provider)
        require_equal(receipt.get("record_exit_code"), 0,
                       "%s perf record exit" % provider)
        require_equal(receipt.get("binary_sha256"), protocol.get("binary_sha256"),
                      "%s binary hash" % provider)
        require_equal(receipt.get("protocol_sha256"), protocol_sha,
                      "%s protocol hash" % provider)
        require_equal(receipt.get("driver_sha256"), capture_driver_sha,
                      "%s capture driver hash" % provider)
        require_equal(receipt.get("verifier_sha256"), verifier_sha,
                      "%s verifier hash" % provider)
        argv = receipt.get("argv")
        ensure(isinstance(argv, list), "%s receipt argv is not a list" % provider)
        ensure("perf" in argv and "--call-graph" in argv and
               argv[argv.index("--call-graph") + 1] == "fp",
               "%s receipt is not an fp perf capture" % provider)
        require_equal(argv[argv.index("--provider") + 1], provider,
                      "%s argv provider" % provider)
        require_equal(argv[argv.index("--corpus") + 1], protocol.get("corpus"),
                      "%s argv corpus" % provider)

        artifacts = receipt.get("artifacts")
        ensure(isinstance(artifacts, dict), "%s receipt artifacts missing" % provider)
        expected_artifacts = {
            "profiles/%s-record.log" % provider,
            "profiles/%s-report.log" % provider,
            "profiles/%s-script.log" % provider,
            "profiles/%s.data" % provider,
            "profiles/%s.json" % provider,
        }
        require_equal(set(artifacts), expected_artifacts,
                      "%s receipt artifact names" % provider)
        verified_artifacts = {}
        compression = None
        compression_path = root / "compression.json"
        if compression_path.is_file():
            compression = read_raw_json(root, Path("compression.json"))
        for artifact_name in sorted(artifacts):
            verified_artifacts[artifact_name] = verify_logical_artifact(
                root, Path(artifact_name), artifacts[artifact_name], compression)

        report_relative = Path("profiles/%s.json" % provider)
        report = read_raw_json(root, report_relative)
        require_equal(report.get("schema"), "pptx_provider_lifecycle_v1",
                      "%s report schema" % provider)
        require_equal(report.get("provider"), provider, "%s report provider" % provider)
        require_equal(report.get("corpus"), protocol.get("corpus"),
                      "%s report corpus" % provider)
        require_equal(report.get("samples"), protocol.get("samples"),
                      "%s report sample count" % provider)
        require_equal(report.get("warmup"), protocol.get("warmup"),
                      "%s report warmup" % provider)
        require_equal(report.get("source_revision"), protocol.get("source_revision"),
                      "%s report source revision" % provider)
        require_equal(report.get("binary_sha256"), protocol.get("binary_sha256"),
                      "%s report binary hash" % provider)
        receipts.append({
            "provider": provider,
            "path": receipt_relative_text,
            "sha256": sha256_bytes(receipt_bytes),
            "bytes": len(receipt_bytes),
            "artifacts": verified_artifacts,
            "report": {
                "schema": report.get("schema"),
                "corpus": report.get("corpus"),
                "samples": report.get("samples"),
                "warmup": report.get("warmup"),
                "checked_iteration_count": report.get("checked_iteration_count"),
                "source_revision": report.get("source_revision"),
                "binary_sha256": report.get("binary_sha256"),
            },
        })
    require_equal(sorted(seen_providers), sorted(EXPECTED_PROVIDERS),
                  "receipt providers")
    return protocol, {
        "protocol": {"path": str(protocol_relative), "sha256": protocol_sha,
                     "bytes": len(protocol_bytes)},
        "profile_index": {"path": str(index_relative), "sha256": index_sha,
                          "bytes": len(index_bytes)},
        "receipts": sorted(receipts, key=lambda row: row["provider"]),
    }


def validate_source_audit(root, protocol):
    index = read_raw_json(root, Path("profile-index.json"))
    first_receipt = read_raw_json(root, Path(index[0]))
    source = first_receipt.get("source")
    ensure(isinstance(source, dict), "receipt source custody is missing")
    for receipt_path in index:
        receipt = read_raw_json(root, Path(receipt_path))
        require_equal(receipt.get("source"), source, "shared source custody")

    source_relative = Path(source["path"])
    source_bytes = raw_json_bytes(root, source_relative)
    require_equal(sha256_bytes(source_bytes), source.get("sha256"),
                  "retained source manifest hash")
    require_equal(source_relative.name, source.get("sha256") + ".json",
                  "retained source manifest name")
    source_manifest = read_raw_json(root, source_relative)
    ensure(isinstance(source_manifest, dict), "source manifest is not an object")
    require_equal(len(source_manifest), source.get("files"), "source manifest file count")

    input_relative = Path("input-source-manifest.json")
    input_bytes = raw_json_bytes(root, input_relative)
    require_equal(input_bytes, source_bytes, "input and retained source manifests")
    input_manifest = read_raw_json(root, input_relative)
    require_equal(input_manifest, source_manifest, "input and retained source manifest JSON")

    input_build_relative = Path("input-build-0429.json")
    input_build_bytes = raw_json_bytes(root, input_build_relative)
    input_build = read_raw_json(root, input_build_relative)
    require_equal(input_build.get("revision"), protocol.get("source_revision"),
                  "input build source revision")
    require_equal(input_build.get("binary_sha256"), protocol.get("binary_sha256"),
                  "input build binary hash")
    require_equal(input_build.get("source_manifest"), source,
                  "input build source manifest")

    check_relative = Path("checks/frame-pointer-profiles.json")
    check_bytes = raw_json_bytes(root, check_relative)
    check = read_raw_json(root, check_relative)
    require_equal(check.get("status"), "pass", "frame-pointer capture check status")
    require_equal(check.get("exit_code"), 0, "frame-pointer capture exit")
    require_equal(check.get("source_unchanged"), True,
                  "frame-pointer source unchanged flag")
    require_equal(check.get("driver_sha256"), sha256_path(root / "check.py"),
                  "frame-pointer custody-check driver hash")
    require_equal(check.get("source_before"), source,
                  "frame-pointer source_before")
    require_equal(check.get("source_after"), source,
                  "frame-pointer source_after")
    check_log = check.get("log")
    ensure(isinstance(check_log, dict), "frame-pointer capture check log is missing")
    compression_path = root / "compression.json"
    compression = read_raw_json(root, Path("compression.json")) \
        if compression_path.is_file() else None
    check_log_binding = verify_logical_artifact(
        root, Path(check_log["path"]), check_log, compression)

    symbol_relative = Path("checks/symbol-identity.json")
    symbol_bytes = raw_json_bytes(root, symbol_relative)
    symbol_check = read_raw_json(root, symbol_relative)
    require_equal(symbol_check.get("status"), "pass", "symbol identity check status")
    require_equal(symbol_check.get("exit_code"), 0, "symbol identity check exit")
    require_equal(symbol_check.get("source_unchanged"), True,
                  "symbol identity source unchanged flag")
    require_equal(symbol_check.get("driver_sha256"), sha256_path(root / "check.py"),
                  "symbol identity custody-check driver hash")
    require_equal(symbol_check.get("source_before"), source,
                  "symbol identity source_before")
    require_equal(symbol_check.get("source_after"), source,
                  "symbol identity source_after")
    symbol_log = symbol_check.get("log")
    ensure(isinstance(symbol_log, dict), "symbol identity log is missing")
    symbol_log_binding = verify_logical_artifact(
        root, Path(symbol_log["path"]), symbol_log, compression)
    symbol_lines = []
    with open_logical(root, Path(symbol_log["path"])) as stream:
        for raw_line in stream:
            line = raw_line.decode("utf-8", "replace").rstrip("\r\n")
            if "::run_lifecycle_iteration" in line:
                symbol_lines.append(line)
    require_equal(len(symbol_lines), 1,
                  "unique run_lifecycle_iteration bound-symbol count")
    ensure("litchi_perf_baseline::pptx_provider_lifecycle::run_lifecycle_iteration" in
           symbol_lines[0], "run_lifecycle_iteration symbol owner")

    return {
        "source_manifest": {
            "path": str(source_relative),
            "sha256": source["sha256"],
            "bytes": len(source_bytes),
            "files": source["files"],
        },
        "input_source_manifest": {
            "path": str(input_relative),
            "sha256": sha256_bytes(input_bytes),
            "bytes": len(input_bytes),
        },
        "input_build": {
            "path": str(input_build_relative),
            "sha256": sha256_bytes(input_build_bytes),
            "bytes": len(input_build_bytes),
            "revision": input_build.get("revision"),
            "binary_sha256": input_build.get("binary_sha256"),
        },
        "capture_check": {
            "path": str(check_relative),
            "sha256": sha256_bytes(check_bytes),
            "bytes": len(check_bytes),
            "status": check.get("status"),
            "source_unchanged": check.get("source_unchanged"),
            "log": check_log_binding,
        },
        "symbol_identity_check": {
            "path": str(symbol_relative),
            "sha256": sha256_bytes(symbol_bytes),
            "bytes": len(symbol_bytes),
            "log": symbol_log_binding,
            "unique_run_lifecycle_iteration_symbols": len(symbol_lines),
            "bound_symbol": symbol_lines[0],
        },
    }


def parse_record_sample_count(root, relative):
    text = read_logical(root, relative).decode("utf-8", "replace")
    matches = RECORD_SAMPLES_RE.findall(text)
    ensure(len(matches) == 1, "expected one perf record sample count in %s" % relative)
    return int(matches[0])


def count_diagnostic_lines(root, relative):
    lines = 0
    kinds = Counter()
    examples = []
    with open_logical(root, relative) as stream:
        for raw_line in stream:
            line = raw_line.decode("utf-8", "replace").rstrip("\r\n")
            lowered = line.casefold()
            if "addr2line" not in lowered and "could not read first record" not in lowered:
                continue
            lines += 1
            if "addr2line" in lowered:
                kinds["addr2line"] += 1
            if "could not read first record" in lowered:
                kinds["could_not_read_first_record"] += 1
            if len(examples) < 4:
                examples.append(line.strip())
    return {"lines": lines, "kinds": dict(kinds), "examples": examples}


def classify_stack(symbols):
    has_iteration = any(ITERATION_RE.search(symbol) for symbol in symbols)
    has_setup = any(marker in symbol for symbol in symbols for marker in SETUP_MARKERS)
    if has_iteration and not has_setup:
        return "iteration", has_iteration, has_setup
    if has_setup and not has_iteration:
        return "corpus_setup", has_iteration, has_setup
    # A stack carrying both ancestry markers is ambiguous.  Keep it in the
    # conservative other bucket and expose the ambiguity separately.
    return "other_or_unresolved", has_iteration, has_setup


def parse_perf_script(root, relative, record_relative):
    categories = Counter({name: 0 for name in CATEGORY_NAMES})
    leaves = {name: Counter() for name in CATEGORY_NAMES}
    api_counts = Counter({name: 0 for name in API_MARKERS})
    deflate_categories = Counter({name: 0 for name in CATEGORY_NAMES})
    diagnostics = []
    diagnostic_lines = 0
    diagnostic_kinds = Counter()
    malformed_frame_lines = 0
    malformed_blocks = 0
    ambiguous_blocks = 0
    decoded_blocks = 0
    total_blocks = 0
    counting_sink_deflate = 0
    counting_sink_deflate_iteration = 0
    source_reader_counting_sink_deflate_iteration = 0
    publication_counting_sink_deflate_iteration = 0
    deflate_iteration = 0
    deflate_setup = 0

    def consume(symbols):
        nonlocal malformed_blocks, ambiguous_blocks, decoded_blocks
        nonlocal total_blocks, counting_sink_deflate
        nonlocal counting_sink_deflate_iteration
        nonlocal source_reader_counting_sink_deflate_iteration
        nonlocal publication_counting_sink_deflate_iteration
        nonlocal deflate_iteration, deflate_setup
        total_blocks += 1
        if symbols:
            decoded_blocks += 1
        else:
            malformed_blocks += 1
        category, has_iteration, has_setup = classify_stack(symbols)
        if has_iteration and has_setup:
            ambiguous_blocks += 1
        categories[category] += 1
        leaves[category][symbols[0] if symbols else "[no decoded frame]"] += 1
        stack = "\n".join(symbols)
        # DeflateDecoder/read-side frames also contain the word "deflate".
        # The retained evidence metric deliberately names the exact compressor
        # leaf used by the publication path.
        is_deflate = bool(DEFLATE_MEDIUM_RE.search(stack))
        is_counting_sink = "CountingSink" in stack
        is_source_reader = "SourceReader" in stack
        is_publication = "publish_cross_slide_copy_to_stream" in stack
        if is_deflate:
            deflate_categories[category] += 1
            if category == "iteration":
                deflate_iteration += 1
            elif category == "corpus_setup":
                deflate_setup += 1
        if is_deflate and is_counting_sink:
            counting_sink_deflate += 1
            if category == "iteration":
                counting_sink_deflate_iteration += 1
                if is_source_reader:
                    source_reader_counting_sink_deflate_iteration += 1
                if is_publication:
                    publication_counting_sink_deflate_iteration += 1
        if category == "iteration":
            for name, marker in API_MARKERS.items():
                if marker in stack:
                    api_counts[name] += 1

    in_sample = False
    symbols = []
    try:
        with open_logical(root, relative) as stream:
            for raw_line in stream:
                line = raw_line.decode("utf-8", "replace").rstrip("\r\n")
                if DIAGNOSTIC_RE.search(line):
                    diagnostic_lines += 1
                    lowered = line.casefold()
                    if "addr2line" in lowered:
                        diagnostic_kinds["addr2line"] += 1
                    if "could not read first record" in lowered:
                        diagnostic_kinds["could_not_read_first_record"] += 1
                    if len(diagnostics) < 12:
                        diagnostics.append(line.strip())
                if SAMPLE_HEADER_RE.match(line):
                    if in_sample:
                        consume(symbols)
                    in_sample = True
                    symbols = []
                    continue
                if not in_sample:
                    continue
                if not line.strip():
                    consume(symbols)
                    in_sample = False
                    symbols = []
                    continue
                frame = FRAME_RE.match(line)
                if frame:
                    symbol = frame.group(1).strip()
                    if symbol:
                        symbols.append(symbol)
                    else:
                        malformed_frame_lines += 1
                elif line[:1].isspace():
                    # A stack-looking line which has no address is retained as
                    # a parse defect, never guessed into an ancestry bucket.
                    malformed_frame_lines += 1
        if in_sample:
            consume(symbols)
    except (OSError, UnicodeError) as error:
        fail("cannot parse %s: %s" % (relative, error))

    recorded_samples = parse_record_sample_count(root, record_relative)
    require_equal(total_blocks, recorded_samples,
                  "%s decoded sample blocks versus perf record" % relative)
    ensure(total_blocks > 0, "no perf sample blocks parsed from %s" % relative)
    total = total_blocks
    percentages = {name: categories[name] * 100.0 / total for name in CATEGORY_NAMES}
    resolved = categories["iteration"] + categories["corpus_setup"]
    record_diagnostics = count_diagnostic_lines(root, record_relative)
    report_relative = Path(str(relative).replace("-script.log", "-report.log"))
    report_diagnostics = count_diagnostic_lines(root, report_relative)
    return {
        "samples": total,
        "recorded_samples": recorded_samples,
        "exclusive_ancestry_counts": dict(categories),
        "exclusive_ancestry_percent": percentages,
        "ancestry_coverage": {
            "resolved_iteration_or_setup_samples": resolved,
            "resolved_iteration_or_setup_percent": resolved * 100.0 / total,
            "unresolved_or_ambiguous_samples": categories["other_or_unresolved"],
            "unresolved_or_ambiguous_percent":
                categories["other_or_unresolved"] * 100.0 / total,
        },
        "iteration_api_ancestry_counts_nonexclusive": dict(api_counts),
        "deflate_intersections": {
            "marker": "deflate_medium",
            "deflate_medium_samples": sum(deflate_categories.values()),
            "deflate_medium_by_exclusive_ancestry": dict(deflate_categories),
            "deflate_medium_iteration_samples": deflate_iteration,
            "deflate_medium_corpus_setup_samples": deflate_setup,
            "counting_sink_deflate_medium_samples": counting_sink_deflate,
            "counting_sink_deflate_medium_iteration_samples": counting_sink_deflate_iteration,
            "source_reader_counting_sink_deflate_medium_iteration_samples":
                source_reader_counting_sink_deflate_iteration,
            "publication_counting_sink_deflate_medium_iteration_samples":
                publication_counting_sink_deflate_iteration,
        },
        "top_leaf_symbols_by_ancestry": {
            name: [[symbol, count] for symbol, count in leaves[name].most_common(20)]
            for name in CATEGORY_NAMES
        },
        "parse_quality": {
            "sample_blocks": total,
            "blocks_with_decoded_frames": decoded_blocks,
            "blocks_without_decoded_frames": malformed_blocks,
            "malformed_frame_lines": malformed_frame_lines,
            "ambiguous_ancestry_blocks": ambiguous_blocks,
            "diagnostic_lines": diagnostic_lines,
            "diagnostic_kinds": dict(diagnostic_kinds),
            "diagnostic_examples": diagnostics,
            "diagnostics_by_artifact": {
                "script": {
                    "lines": diagnostic_lines,
                    "kinds": dict(diagnostic_kinds),
                },
                "record": record_diagnostics,
                "report": report_diagnostics,
            },
            "diagnostic_rule": "retain addr2line and could-not-read-first-record lines; do not infer frames",
        },
        "stack_input": {
            "path": str(relative),
            "accepted_storage": "raw or gzip; receipt raw hash and byte count are checked",
        },
        "scope_note": (
            "Iteration ancestry is a stack marker over the three warmups and 100 "
            "retained iterations; perf stacks do not identify warmup versus retained "
            "sample. Counts are unweighted full-process cycles:u sample blocks, not "
            "API wall-time fractions."
        ),
    }


def historical_comparison(root, profiles):
    relative = Path("input-profile-summary-0429.json")
    path = root / relative
    if not path.is_file():
        return {"available": False, "reason": "historical 0429 summary is not in this bundle"}
    try:
        historical = json.loads(path.read_text(), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        return {"available": False, "reason": "historical summary unreadable: %s" % error}
    by_name = {row.get("name"): row for row in historical.get("profiles", [])}
    rows = []
    for current in profiles:
        old_name = "media-rich-" + current["provider"]
        old = by_name.get(old_name)
        if not isinstance(old, dict):
            rows.append({"provider": current["provider"], "historical_profile": old_name,
                         "available": False})
            continue
        current_counts = current["exclusive_ancestry_counts"]
        old_counts = old.get("exclusive_ancestry_counts", {})
        rows.append({
            "provider": current["provider"],
            "historical_profile": old_name,
            "available": True,
            "current_samples": current["samples"],
            "historical_samples": old.get("samples"),
            "current_exclusive_ancestry_counts": current_counts,
            "historical_exclusive_ancestry_counts": old_counts,
        })
    return {
        "available": True,
        "path": str(relative),
        "sha256": sha256_path(path),
        "profiles": rows,
        "interpretation": (
            "Descriptive sample-count comparison only. The historical profiles use "
            "the retained DWARF16384 call graph while this bundle uses fp; neither "
            "set of counts is an API wall-time fraction or a speedup estimate."
        ),
    }


def frozen_custody_expectation(root):
    """Load only custody pins before parsing large retained stacks, if frozen."""
    target = root / SUMMARY_NAME
    if not target.is_file():
        return None
    try:
        summary = json.loads(target.read_text(), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail("invalid %s: %s" % (target, error))
    ensure(isinstance(summary, dict), "%s is not an object" % target)
    return {
        "artifact_binding": summary.get("artifact_binding"),
        "source_audit": summary.get("source_audit"),
    }


def derive(root):
    frozen = frozen_custody_expectation(root)
    protocol, binding = validate_protocol_and_receipts(root)
    if frozen is not None:
        require_equal(binding, frozen["artifact_binding"],
                      "frozen artifact custody binding")
    source_audit = validate_source_audit(root, protocol)
    if frozen is not None:
        require_equal(source_audit, frozen["source_audit"],
                      "frozen source audit binding")
    profiles = []
    for receipt in binding["receipts"]:
        provider = receipt["provider"]
        script_relative = Path("profiles/%s-script.log" % provider)
        record_relative = Path("profiles/%s-record.log" % provider)
        parsed = parse_perf_script(root, script_relative, record_relative)
        parsed.update({
            "provider": provider,
            "receipt": receipt["path"],
            "receipt_sha256": receipt["sha256"],
            "artifact_binding": receipt["artifacts"],
        })
        profiles.append(parsed)
    profiles.sort(key=lambda row: row["provider"])
    return {
        "status": "pass",
        "schema": "pptx_frame_pointer_profile_analysis_v1",
        "change": 430,
        "protocol": {
            "path": "profile-protocol.json",
            "sha256": binding["protocol"]["sha256"],
            "binary_sha256": protocol.get("binary_sha256"),
            "source_revision": protocol.get("source_revision"),
            "event": protocol.get("event"),
            "call_graph": protocol.get("call_graph"),
            "frequency_hz": protocol.get("frequency_hz"),
            "corpus": protocol.get("corpus"),
            "providers": protocol.get("providers"),
            "samples": protocol.get("samples"),
            "warmup": protocol.get("warmup"),
        },
        "artifact_binding": binding,
        "source_audit": source_audit,
        "profiles": profiles,
        "scope": (
            "Evidence-only interpretation of retained raw or gzip perf script stacks "
            "from the unchanged 0429 binary. Iteration ancestry includes warmup and "
            "retained sample iterations because the stacks do not carry the iteration "
            "index. API names are nonexclusive. Exact deflate_medium and CountingSink "
            "intersections are stack co-occurrence evidence, not measured API timing."
        ),
        "historical_0429_comparison": historical_comparison(root, profiles),
        "performance_claim": None,
        "analysis_script_sha256": sha256_path(root / Path(__file__).name),
    }


def compare_summary(root, result):
    target = root / SUMMARY_NAME
    ensure(target.is_file(), "missing %s" % target)
    try:
        expected = json.loads(target.read_text(), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail("invalid %s: %s" % (target, error))
    ensure(expected == json.loads(json.dumps(result)),
           "%s does not match retained evidence" % target)


def run_child_check(bundle):
    command = [sys.executable, "-B", str(bundle / Path(__file__).name), "--check"]
    return subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                          check=False)


def mutate_append(path):
    original_size = path.stat().st_size
    with path.open("ab") as stream:
        stream.write(b"\n")

    def restore():
        with path.open("r+b") as stream:
            stream.truncate(original_size)
    return restore


def mutate_one_byte(path):
    offset = max(0, path.stat().st_size // 2)
    with path.open("r+b") as stream:
        stream.seek(offset)
        original = stream.read(1)
        ensure(original, "cannot mutate empty artifact %s" % path)
        stream.seek(offset)
        stream.write(bytes([original[0] ^ 1]))

    def restore():
        with path.open("r+b") as stream:
            stream.seek(offset)
            stream.write(original)
    return restore


def portable_check(root):
    """Replay --check from a copied bundle and prove custody mutations fail."""
    with tempfile.TemporaryDirectory(prefix="litchi-0430-profile-replay-") as directory:
        parent = Path(directory)
        bundle = parent / "bundle"
        shutil.copytree(root, bundle)
        result = run_child_check(bundle)
        ensure(result.returncode == 0, "portable copied-bundle --check failed")

        mutation_paths = [
            ("protocol", bundle / "profile-protocol.json", mutate_append),
            ("profile-index", bundle / "profile-index.json", mutate_append),
            ("receipt", bundle / "profiles/bytes-receipt.json", mutate_append),
            ("source-audit", bundle / "checks/frame-pointer-profiles.json", mutate_append),
        ]
        stack_path = bundle / "profiles/bytes-script.log"
        if not stack_path.is_file():
            stack_path = bundle / "profiles/bytes-script.log.gz"
        mutation_paths.append(("retained-stack", stack_path, mutate_one_byte))
        for label, path, mutate in mutation_paths:
            ensure(path.is_file(), "portable mutation target missing: %s" % path)
            restore = mutate(path)
            try:
                result = run_child_check(bundle)
                ensure(result.returncode != 0,
                       "portable mutation unexpectedly accepted: %s" % label)
            finally:
                restore()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="derive and compare against the frozen profile-summary.json")
    parser.add_argument("--portable-check", action="store_true",
                        help="run --check from a copied bundle and exercise custody mutations")
    args = parser.parse_args()
    try:
        result = derive(ROOT)
        if args.check or args.portable_check:
            compare_summary(ROOT, result)
        else:
            target = ROOT / SUMMARY_NAME
            ensure(not target.exists(), "%s already exists; use --check" % target)
            target.write_text(json.dumps(result, indent=2) + "\n")
        if args.portable_check:
            portable_check(ROOT)
        print(json.dumps({"status": "pass", "profiles": len(result["profiles"]),
                          "portable": args.portable_check}))
    except AnalysisError as error:
        print("profile-analysis: %s" % error, file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
