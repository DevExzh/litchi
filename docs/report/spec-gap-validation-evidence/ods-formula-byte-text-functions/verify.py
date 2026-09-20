#!/usr/bin/env python3
"""Independently verify the frozen byte-position evaluator evidence.

This verifier is deliberately fail-closed.  It checks the gate receipt,
re-runs the independent oracle, parses the retained native ODS itself, and
requires complete performance captures before it reports success.  Performance
rows, raw checksums, and source manifests are checked directly; PASS markers
and summary prose are not treated as evidence.
"""

from __future__ import annotations

import ast
from collections import defaultdict
import hashlib
import json
from pathlib import Path
import re
import statistics
import subprocess
import zipfile
import xml.etree.ElementTree as ET


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
GATES = HERE / "gates"
PERFORMANCE = HERE / "performance"
RESULTS = PERFORMANCE / "results"
BASELINE_JSON = HERE / "baseline.json"

FUNCTIONS = ("FINDB", "LEFTB", "LENB", "MIDB", "REPLACEB", "RIGHTB", "SEARCHB")
BASELINE_COMMIT = json.loads(BASELINE_JSON.read_text(encoding="utf-8"))["commit"]
PHASES = ("evaluate", "parse-evaluate")
SAMPLES = 15
WARMUPS = 3
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"

CONTROL_CASES = (
    "scalar-control-arithmetic",
    "scalar-control-sin",
    "scalar-control-imsum",
    "database-control-dsum",
    "scalar-control-average",
    "scalar-control-counta",
    "scalar-control-var",
    "scalar-control-stdev",
    "database-control-dvar",
    "database-control-dstdev",
    "array-control-4x4-arithmetic",
    "array-control-4x4-sin",
    "array-control-16x16-arithmetic",
    "array-control-16x16-sin",
    "reference-array-16x4-arithmetic",
    "scalar-aggregate-sum",
    "literal-aggregate-4x1-sum",
    "reference-aggregate-64x4-sum",
    "reference-conditional-256x4-sumifs",
    "reference-control-average",
    "reference-control-counta",
    "representative-median",
    "representative-rank",
    "representative-percentrank",
    "concat-borrowed-literals",
    "concat-owned-left",
    "concat-owned-right",
    "concat-growth-chain",
)
BYTE_CASES = tuple(
    [
        f"{lane}-byte-{function}"
        for lane in ("tiny", "large-unicode", "large-ascii", "reference-64", "refusal")
        for function in ("findb", "leftb", "lenb", "midb", "replaceb", "rightb", "searchb")
    ]
    + [f"matrix-broadcast-byte-{function}" for function in ("findb", "leftb", "lenb", "midb", "replaceb", "rightb", "searchb")]
    + [f"cancellation-byte-{function}" for function in ("lenb", "midb", "replaceb", "searchb")]
    + [f"resource-byte-{function}" for function in ("lenb", "midb", "replaceb", "searchb")]
    + [f"search-worstcase-byte-{function}" for function in ("findb", "searchb")]
    + ["replaceb-growth"]
    + [
        "projected-statistical-average-lenb",
        "projected-statistical-sum-lenb",
        "projected-statistical-average-len",
    ]
)
EXPECTED_CANDIDATE_CASES = CONTROL_CASES + BYTE_CASES
BASELINE_LABEL = f"baseline-{BASELINE_COMMIT}"


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load(path: Path):
    if not path.is_file():
        raise RuntimeError(f"missing evidence file: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def equal(label: str, observed, expected) -> None:
    if observed != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, observed {observed!r}")


def verify_gates() -> dict:
    path = GATES / "verify.py"
    run = subprocess.run(
        ["python3", str(path)], cwd=REPO, capture_output=True, text=True
    )
    if run.returncode:
        raise RuntimeError(
            f"gates/verify.py failed ({run.returncode}):\n{run.stdout}{run.stderr}"
        )
    try:
        receipt = json.loads(run.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"gates/verify.py did not emit JSON: {run.stdout!r}") from error
    equal("gate verification flag", receipt.get("verified"), True)
    equal("gate count", receipt.get("gates"), 7)
    equal("gate focused evaluation count", receipt.get("focused", {}).get("evaluation"), 6)
    equal("gate focused limits count", receipt.get("focused", {}).get("limits"), 5)
    equal("gate focused oracle count", receipt.get("focused", {}).get("oracle"), 2)
    return receipt


def verify_oracle() -> dict:
    contract = HERE / "contract.md"
    goldens = load(HERE / "byte-goldens.json")
    equal("oracle contract hash", goldens.get("contract_sha256"), digest(contract))
    rows = goldens.get("observations")
    if not isinstance(rows, list):
        raise RuntimeError("byte-goldens.json observations is not a list")
    equal("oracle function set", sorted({row.get("function") for row in rows}), sorted(FUNCTIONS))
    run = subprocess.run(
        ["python3", str(HERE / "byte_oracle.py"), "--check"],
        cwd=HERE,
        capture_output=True,
        text=True,
    )
    if run.returncode:
        raise RuntimeError(f"byte_oracle.py --check failed:\n{run.stdout}{run.stderr}")
    receipt = json.loads(run.stdout)
    equal("oracle receipt", receipt, {"functions": 7, "observations": len(rows), "verified": True})
    equal("oracle observation count", len(rows), 1376)
    return receipt


TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
STRING_VALUE = OFFICE + "string-value"


def cell_text(cell: ET.Element) -> str:
    return "".join(cell.itertext())


def typed_cell(cell: ET.Element) -> dict[str, object]:
    kind = cell.attrib.get(VALUE_TYPE)
    if kind == "float":
        raw = cell.attrib[VALUE]
        value: object = int(raw) if re.fullmatch(r"[-+]?\d+", raw) else float(raw)
        return {"type": "number", "value": value}
    if kind == "string":
        return {"type": "text", "value": cell.attrib.get(STRING_VALUE, cell_text(cell))}
    if kind == "error":
        return {"type": "error", "value": cell.attrib.get(VALUE, cell_text(cell))}
    raise RuntimeError(f"unexpected native cell type: {kind!r}")


def formula_rows(xml_bytes: bytes, *, typed: bool) -> list[dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    rows: list[dict[str, object]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formulas = [cell for cell in cells if FORMULA in cell.attrib]
        if not formulas:
            continue
        if len(cells) != 3 or len(formulas) != 1:
            raise RuntimeError("native fixture row shape changed")
        formula_cell = formulas[0]
        item: dict[str, object] = {
            "case": cell_text(cells[0]),
            "formula": formula_cell.attrib[FORMULA],
        }
        if typed:
            item["native"] = typed_cell(formula_cell)
        else:
            item["expected_display"] = cell_text(cells[2])
        rows.append(item)
    return rows


def verify_native() -> dict:
    directory = HERE / "native"
    provenance = load(directory / "provenance.json")
    fixture = provenance.get("fixture", {})
    input_path = directory / str(fixture.get("input", ""))
    output_path = directory / str(fixture.get("recalculated_output", ""))
    if not input_path.is_file() or not output_path.is_file():
        raise RuntimeError("native fixture input/output is incomplete")
    equal("native input hash", digest(input_path), fixture.get("input_sha256"))
    equal("native ODS hash", digest(output_path), fixture.get("recalculated_output_sha256"))
    with zipfile.ZipFile(output_path) as archive:
        content = archive.read("content.xml")
    equal("native content.xml hash", digest_bytes(content), fixture.get("recalculated_content_xml_sha256"))
    with zipfile.ZipFile(output_path) as archive:
        observed = formula_rows(archive.read("content.xml"), typed=True)
    with input_path.open("rb") as stream:
        source = formula_rows(stream.read(), typed=False)
    expected = load(directory / "native-results.json")
    equal("native source row count", len(source), 19)
    equal("native output row count", len(observed), 19)
    equal(
        "native formula sequence",
        [(row["case"], row["formula"]) for row in source],
        [(row["case"], row["formula"]) for row in expected],
    )
    expected_projection = [
        {"case": row["case"], "formula": row["formula"], "native": row["native"]}
        for row in expected
    ]
    equal("native typed rows", observed, expected_projection)
    for row in expected:
        equal(
            f"native comparison flag {row['case']}",
            row.get("matches_profile"),
            row.get("native") == row.get("profile"),
        )
    equal("native function set", sorted({row["function"] for row in expected}), sorted(FUNCTIONS))
    ascii_rows = [row for row in expected if str(row["case"]).startswith("ascii_")]
    nonascii_rows = [row for row in expected if str(row["case"]).startswith("utf8_")]
    equal("native ASCII rows", len(ascii_rows), 7)
    equal("native ASCII parity", sum(row["matches_profile"] for row in ascii_rows), 7)
    equal("native non-ASCII rows", len(nonascii_rows), 12)
    equal("native non-ASCII divergences", sum(not row["matches_profile"] for row in nonascii_rows), 10)
    return {
        "rows": len(expected),
        "functions": len(FUNCTIONS),
        "ascii_matching": len(ascii_rows),
        "nonascii_divergences": sum(not row["matches_profile"] for row in nonascii_rows),
    }


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def profile_snapshot() -> dict[str, str]:
    excluded = {"results", "target", "__pycache__"}
    result: dict[str, str] = {}
    for path in sorted(PERFORMANCE.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(PERFORMANCE)
        if any(part in excluded for part in relative.parts):
            continue
        result[str(relative)] = digest(path)
    if not result:
        raise RuntimeError("performance profile input directory is empty")
    return result


def git_file_digest(commit: str, relative: str) -> str | None:
    try:
        data = subprocess.check_output(
            ["git", "show", f"{commit}:{relative}"],
            cwd=REPO,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        return None
    return digest_bytes(data)


def expected_workspace_paths(commit: str) -> set[str]:
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", commit, "crates"], cwd=REPO, text=True
    )
    for line in listing.splitlines():
        if line.endswith(".rs") or line.endswith("Cargo.toml") or line.endswith("build.rs"):
            paths.add(line)
    return paths


def fixed_profile_paths() -> set[str]:
    """Read only the runner's explicit source list for baseline key custody."""
    tree = ast.parse((PERFORMANCE / "run_profile.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if not any(isinstance(target, ast.Name) and target.id == "SOURCE_FILES" for target in node.targets):
            continue
        value = ast.literal_eval(node.value)
        if not isinstance(value, tuple):
            raise RuntimeError("performance SOURCE_FILES is not a tuple")
        return set(value)
    raise RuntimeError("performance SOURCE_FILES declaration is missing")


def verify_source_map(
    label: str,
    manifest: dict,
    freeze: dict,
    *,
    candidate: bool,
) -> None:
    before = manifest.get("before")
    after = manifest.get("after")
    if not isinstance(before, dict) or not isinstance(after, dict):
        raise RuntimeError(f"{label}: malformed source manifest")
    for key in ("source_sha256", "workspace_source_sha256", "profile_input_sha256"):
        equal(f"{label} {key} stable", before.get(key), after.get(key))
    equal(f"{label} selected sources stable", manifest.get("source_sha256_unchanged"), True)
    equal(f"{label} workspace sources stable", manifest.get("workspace_source_sha256_unchanged"), True)
    equal(f"{label} profile stable", manifest.get("profile_input_sha256_unchanged"), True)
    equal(f"{label} git stable", manifest.get("git_head_unchanged"), True)
    equal(f"{label} profile inputs", before.get("profile_input_sha256"), profile_snapshot())
    expected_lock = digest(GATES / "Cargo.lock")
    equal(f"{label} frozen workspace lock", before.get("workspace_lock_sha256"), expected_lock)
    equal(f"{label} gate lock hash", expected_lock, GATE_LOCK_SHA256)

    selected = freeze["selected_files"]
    selected_map = before.get("source_sha256", {})
    if not isinstance(selected_map, dict):
        raise RuntimeError(f"{label}: selected source map is malformed")
    staged_profile = load(GATES / "staged-profile-sources.json")
    if not isinstance(staged_profile, dict):
        raise RuntimeError("staged profile source map is malformed")
    for relative, expected in selected.items():
        equal(f"frozen selected/staged source {relative}", staged_profile.get(relative), expected)
    if candidate:
        equal(f"{label} full staged source path set", set(selected_map), set(staged_profile))
    else:
        fixed = fixed_profile_paths()
        expected_baseline_paths = {
            relative
            for relative in staged_profile
            if git_file_digest(freeze["base_commit"], relative) is not None
            or relative in fixed
        }
        equal(f"{label} full baseline source path set", set(selected_map), expected_baseline_paths)
    for relative, observed in selected_map.items():
        if relative == "Cargo.lock":
            expected = expected_lock
        elif candidate and relative in selected:
            expected = staged_profile[relative]
        elif candidate:
            expected = staged_profile[relative]
        else:
            expected = git_file_digest(freeze["base_commit"], relative)
        equal(f"{label} selected source {relative}", observed, expected)

    workspace_map = before.get("workspace_source_sha256", {})
    if not isinstance(workspace_map, dict):
        raise RuntimeError(f"{label}: workspace source map is malformed")
    expected_paths = expected_workspace_paths(freeze["base_commit"])
    if candidate:
        expected_paths.update(
            relative
            for relative in selected
            if relative.startswith("crates/")
            and (relative.endswith(".rs") or relative.endswith("Cargo.toml") or relative.endswith("build.rs"))
        )
    equal(f"{label} workspace source path set", set(workspace_map), expected_paths)
    for relative, observed in workspace_map.items():
        if relative == "Cargo.lock":
            expected = expected_lock
        elif candidate and relative in selected:
            expected = selected[relative]
        else:
            expected = git_file_digest(freeze["base_commit"], relative)
        equal(f"{label} workspace source {relative}", observed, expected)

    harness = before.get("harness_sha256")
    if not isinstance(harness, dict):
        raise RuntimeError(f"{label}: harness hash map missing")
    for relative in ("Cargo.toml", "Cargo.lock", "src/main.rs"):
        current = PERFORMANCE / "harness" / relative
        equal(f"{label} harness {relative}", harness.get(relative), digest(current))
    cleanup = load(RESULTS / label / "target-cleanup.json")
    equal(f"{label} target cleanup", cleanup.get("removed"), True)


def preflight_read_bound(case: str) -> int:
    if case in {"database-control-dsum", "database-control-dvar", "database-control-dstdev"}:
        return 7
    if case == "reference-array-16x4-arithmetic":
        return 64
    if case in {"reference-aggregate-64x4-sum", "reference-control-average", "reference-control-counta"}:
        return 256
    if case == "reference-conditional-256x4-sumifs":
        return 1792
    if case.startswith(("tiny-", "large-unicode-", "large-ascii-", "refusal-", "search-worstcase-", "replaceb-growth")):
        return 0
    if case.startswith("reference-64-byte-"):
        return 64
    if case.startswith("matrix-broadcast-byte-"):
        return 128 if case.rsplit("-", 1)[-1] in {"leftb", "midb", "replaceb", "rightb"} else 64
    if case.startswith("cancellation-byte-"):
        return 1
    if case.startswith("resource-byte-"):
        return 0
    if case.startswith("projected-statistical-"):
        return {
            "projected-statistical-average-lenb": 2,
            "projected-statistical-sum-lenb": 4,
            "projected-statistical-average-len": 2,
        }[case]
    return 0


def direct_read_bound(case: str, elements: int) -> int | None:
    if case.startswith("cancellation-byte-"):
        return 1
    return preflight_read_bound(case)


def raw_rows(directory: Path) -> list[dict]:
    path = directory / "measurements.jsonl"
    if not path.is_file():
        raise RuntimeError(f"missing performance rows: {path}")
    rows = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip():
            rows.append(json.loads(line))
    return rows


def verify_results_manifest() -> None:
    retained_path = RESULTS / "retained-files.json"
    retained = load(retained_path)
    if not isinstance(retained, dict):
        raise RuntimeError("performance retained-files.json is not a map")
    actual = {
        str(path.relative_to(RESULTS))
        for path in RESULTS.rglob("*")
        if path.is_file() and path != retained_path
    }
    equal("performance retained file set", set(retained), actual)
    for relative, expected in retained.items():
        equal(f"performance retained hash {relative}", digest(RESULTS / relative), expected)


def verify_preflight(directory: Path, expected_cases: tuple[str, ...]) -> None:
    receipt = load(directory / "preflight.json")
    equal(f"{directory} preflight status", receipt.get("status"), "ok")
    equal(f"{directory} preflight cases", sorted(receipt.get("cases", [])), sorted(expected_cases))
    observed = receipt.get("reference_reads")
    if not isinstance(observed, dict):
        raise RuntimeError(f"{directory}: preflight reference reads missing")
    equal(f"{directory} preflight read keys", sorted(observed), sorted(expected_cases))
    for case in expected_cases:
        equal(f"{directory} preflight reads {case}", int(observed[case]), preflight_read_bound(case))
    stdout = directory / str(receipt.get("stdout", ""))
    if not stdout.is_file():
        raise RuntimeError(f"{directory}: preflight stdout missing")
    text = stdout.read_text(encoding="utf-8")
    equal(f"{directory} preflight-ok count", text.count("preflight-ok cases=1 phase=evaluate"), len(expected_cases))
    matches = re.findall(r"^preflight case=(\S+) reference_reads=(\d+)$", text, re.MULTILINE)
    equal(f"{directory} preflight observations", len(matches), len(expected_cases))
    for case, reads in matches:
        equal(f"{directory} preflight line {case}", int(reads), preflight_read_bound(case))


def verify_capture_rows(
    directory: Path,
    rows: list[dict],
    expected_cases: tuple[str, ...],
) -> dict[tuple[str, str], list[dict]]:
    expected_groups = {(case, phase) for case in expected_cases for phase in PHASES}
    equal(f"{directory} row count", len(rows), len(expected_groups) * SAMPLES)
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for row in rows:
        key = (row.get("case"), row.get("phase"))
        if key not in expected_groups:
            raise RuntimeError(f"{directory}: unexpected group {key}")
        groups[key].append(row)
        equal(f"{directory} {key} capture", row.get("capture"), directory.name)
        equal(f"{directory} {key} supported", row.get("supported"), True)
        equal(f"{directory} {key} warmups", row.get("warmups"), WARMUPS)
        equal(f"{directory} {key} iterations", row.get("iterations"), 1)
        equal(f"{directory} {key} source head", row.get("source_git_head"), BASELINE_COMMIT)
        if int(row.get("repeat", 0)) <= 0:
            raise RuntimeError(f"{directory} {key}: invalid repeat")
        sample_list = row.get("samples")
        if not isinstance(sample_list, list) or len(sample_list) != 1:
            raise RuntimeError(f"{directory} {key}: expected one raw sample")
        sample = sample_list[0]
        for field in (
            "elapsed_ns", "alloc_calls", "dealloc_calls", "requested_bytes", "released_bytes",
            "live_before", "live_after", "peak_live_delta", "work", "memory_retained",
            "reference_reads", "output_bytes", "checksum",
        ):
            if int(sample.get(field, -1)) < 0:
                raise RuntimeError(f"{directory} {key}: invalid {field}")
        equal(f"{directory} {key} balanced live bytes", sample["live_before"], sample["live_after"])
        equal(f"{directory} {key} balanced allocation bytes", sample["requested_bytes"], sample["released_bytes"])
        equal(f"{directory} {key} checksum p50", row.get("checksum_p50"), sample["checksum"])
        equal(f"{directory} {key} output p50", row.get("output_bytes_p50"), sample["output_bytes"])
        repeat = int(row["repeat"])
        equal(
            f"{directory} {key} bytes per repeat",
            row.get("bytes_per_repeat_p50"),
            int(row.get("input_bytes", -1)) + sample["output_bytes"] // repeat,
        )
        equal(f"{directory} {key} normalized reads", row.get("reference_reads_per_repeat"), sample["reference_reads"] // repeat)
        if not str(row.get("validation_scope", "")).startswith("one untimed contract fixture oracle"):
            raise RuntimeError(f"{directory} {key}: validation scope is not recorded")
        raw_stdout = directory / str(row.get("raw_stdout", ""))
        raw_time = directory / str(row.get("raw_time", ""))
        if not raw_stdout.is_file() or not raw_time.is_file():
            raise RuntimeError(f"{directory} {key}: raw measurement files missing")
        raw = json.loads(raw_stdout.read_text(encoding="utf-8"))
        for field, value in raw.items():
            if field != "rss_kib":
                equal(f"{directory} {key} raw {field}", row.get(field), value)
        time_text = raw_time.read_text(encoding="utf-8")
        if "Exit status: 0" not in time_text:
            raise RuntimeError(f"{directory} {key}: timed child did not exit successfully")
        match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", time_text)
        if match is None:
            raise RuntimeError(f"{directory} {key}: RSS missing")
        equal(f"{directory} {key} RSS", row.get("rss_kib"), int(match.group(1)))
        expected_reads = direct_read_bound(str(key[0]), int(row.get("elements", 0)))
        if expected_reads is not None:
            if str(key[0]).startswith("cancellation-byte-"):
                equal(f"{directory} {key} cancellation repeat", repeat, 4)
                equal(f"{directory} {key} cancellation reads", sample["reference_reads"], 1)
            else:
                equal(f"{directory} {key} reference reads", sample["reference_reads"], expected_reads * repeat)
    equal(f"{directory} group set", set(groups), expected_groups)
    for key, group in groups.items():
        equal(f"{directory} {key} sample count", len(group), SAMPLES)
        equal(f"{directory} {key} sample indices", sorted(row["sample_index"] for row in group), list(range(1, SAMPLES + 1)))
        checksums = {row["samples"][0]["checksum"] for row in group}
        if len(checksums) != 1:
            raise RuntimeError(f"{directory} {key}: checksum changed across samples")
        binaries = {row.get("binary_sha256") for row in group}
        if len(binaries) != 1:
            raise RuntimeError(f"{directory} {key}: binary changed across samples")
    return groups


def p50(values: list[int]) -> int:
    ordered = sorted(values)
    return ordered[(len(ordered) - 1) * 50 // 100]


def summary_row(group: list[dict]) -> dict:
    first = group[0]
    row = {
        "case": first["case"],
        "operation": first.get("operation"),
        "phase": first["phase"],
        "shape": first.get("shape"),
        "rows": first.get("rows"),
        "columns": first.get("columns"),
        "elements": first.get("elements"),
        "repeat": first.get("repeat"),
        "samples": len(group),
    }
    for source, target in (
        ("elapsed_ns_per_repeat", "time_ns_per_repeat"),
        ("allocator_calls_p50", "allocator_calls"),
        ("requested_bytes_p50", "requested_bytes"),
        ("released_bytes_p50", "released_bytes"),
        ("peak_live_delta_p50", "peak_live_bytes"),
        ("memory_retained_p50", "result_live_budget"),
        ("work_per_repeat", "work_per_repeat"),
        ("reference_reads_per_repeat", "reference_reads"),
        ("rss_kib", "rss_kib"),
    ):
        row[target] = p50([int(record[source]) for record in group if record.get(source) is not None])
    row["checksum"] = p50([int(record["checksum_p50"]) for record in group])
    row["input_bytes"] = first.get("input_bytes")
    row["output_bytes_p50"] = first.get("output_bytes_p50")
    row["bytes_per_repeat_p50"] = first.get("bytes_per_repeat_p50")
    row["expected"] = first.get("expected")
    return row


def verify_report(
    baseline_groups: dict[tuple[str, str], list[dict]],
    candidate_groups: dict[tuple[str, str], list[dict]],
) -> None:
    report = load(RESULTS / "performance-report.json")
    equal("performance report baseline groups", report.get("baseline_groups"), len(baseline_groups))
    equal("performance report candidate groups", report.get("candidate_groups"), len(candidate_groups))
    reported: dict[str, dict[tuple[str, str], dict]] = {BASELINE_LABEL: {}, "candidate-final": {}}
    for pair in report.get("controls", []):
        for side, label in (("baseline", BASELINE_LABEL), ("candidate", "candidate-final")):
            row = pair.get(side)
            if not isinstance(row, dict):
                raise RuntimeError("performance report control pair is malformed")
            key = (row.get("case"), row.get("phase"))
            if key in reported[label]:
                raise RuntimeError(f"duplicate performance report group {label} {key}")
            reported[label][key] = row
    for row in report.get("candidate_only", []):
        key = (row.get("case"), row.get("phase"))
        if key in reported["candidate-final"]:
            raise RuntimeError(f"duplicate candidate-only performance report group {key}")
        reported["candidate-final"][key] = row
    expected = {
        BASELINE_LABEL: {key: summary_row(value) for key, value in baseline_groups.items()},
        "candidate-final": {key: summary_row(value) for key, value in candidate_groups.items()},
    }
    equal("performance report baseline group keys", set(reported[BASELINE_LABEL]), set(expected[BASELINE_LABEL]))
    equal("performance report candidate group keys", set(reported["candidate-final"]), set(expected["candidate-final"]))
    for label in expected:
        for key, row in expected[label].items():
            equal(f"performance report {label} {key}", reported[label][key], row)


def verify_performance() -> dict:
    if not RESULTS.is_dir():
        raise RuntimeError("performance captures are absent; refusing to pass without raw rows")
    case_matrix = load(PERFORMANCE / "case-matrix.json")
    equal("performance case-matrix candidate count", case_matrix.get("candidate_case_count"), 84)
    equal("performance case-matrix control count", case_matrix.get("matched_control_count"), 28)
    equal("performance case-matrix byte count", case_matrix.get("byte_case_count"), 53)
    equal("performance case-matrix reducer count", case_matrix.get("projected_reducer_count"), 3)
    equal("performance case-matrix functions", case_matrix.get("functions"), list(FUNCTIONS))
    equal("performance case-matrix phases", case_matrix.get("phases"), list(PHASES))
    equal(
        "performance case-matrix read bounds",
        case_matrix.get("read_bounds"),
        {
            "tiny-byte-*": 0,
            "large-unicode-byte-*": 0,
            "large-ascii-byte-*": 0,
            "refusal-byte-*": 0,
            "reference-64-byte-*": 64,
            "matrix-broadcast-byte-lenb": 64,
            "matrix-broadcast-byte-findb": 64,
            "matrix-broadcast-byte-searchb": 64,
            "matrix-broadcast-byte-leftb": 128,
            "matrix-broadcast-byte-midb": 128,
            "matrix-broadcast-byte-replaceb": 128,
            "matrix-broadcast-byte-rightb": 128,
            "cancellation-byte-{lenb,midb,replaceb,searchb}": 1,
            "resource-byte-{lenb,midb,replaceb,searchb}": 0,
            "search-worstcase-byte-{findb,searchb}": 0,
            "replaceb-growth": 0,
            "projected-statistical-average-lenb": 2,
            "projected-statistical-sum-lenb": 4,
            "projected-statistical-average-len": 2,
        },
    )
    verify_results_manifest()
    before_inputs = load(RESULTS / "profile-inputs-before.json")
    after_inputs = load(RESULTS / "profile-inputs-after.json")
    equal("performance profile inputs before/after", before_inputs, after_inputs)
    equal("performance current profile inputs", profile_snapshot(), before_inputs)
    freeze = load(GATES / "freeze.json")
    equal("performance freeze base", freeze.get("base_commit"), BASELINE_COMMIT)
    baseline_dir = RESULTS / BASELINE_LABEL
    candidate_dir = RESULTS / "candidate-final"
    baseline_manifest = load(baseline_dir / "source-manifest.json")
    candidate_manifest = load(candidate_dir / "source-manifest.json")
    verify_source_map(BASELINE_LABEL, baseline_manifest, freeze, candidate=False)
    verify_source_map("candidate-final", candidate_manifest, freeze, candidate=True)
    baseline_rows = raw_rows(baseline_dir)
    candidate_rows = raw_rows(candidate_dir)
    baseline_groups = verify_capture_rows(baseline_dir, baseline_rows, CONTROL_CASES)
    candidate_groups = verify_capture_rows(candidate_dir, candidate_rows, EXPECTED_CANDIDATE_CASES)
    equal("performance baseline/candidate harness hashes", baseline_manifest["before"]["harness_sha256"], candidate_manifest["before"]["harness_sha256"])
    equal("performance baseline/candidate workspace lock", baseline_manifest["before"]["workspace_lock_sha256"], candidate_manifest["before"]["workspace_lock_sha256"])
    equal("performance baseline/candidate profile hash", baseline_manifest["before"]["profile_input_sha256"], candidate_manifest["before"]["profile_input_sha256"])
    before_workspace = baseline_manifest["before"]["workspace_source_sha256"]
    after_workspace = candidate_manifest["before"]["workspace_source_sha256"]
    changed = {path for path in set(before_workspace) | set(after_workspace) if before_workspace.get(path) != after_workspace.get(path)}
    expected_changed = {
        path for path in freeze["selected_files"]
        if path.startswith("crates/")
        and path.endswith(".rs")
        and git_file_digest(freeze["base_commit"], path) != freeze["selected_files"][path]
    }
    equal("performance source closure changed paths", changed, expected_changed)
    baseline_env = load(baseline_dir / "environment.json")
    candidate_env = load(candidate_dir / "environment.json")
    equal("performance baseline cases", baseline_env.get("cases"), list(CONTROL_CASES))
    equal("performance candidate cases", candidate_env.get("cases"), list(EXPECTED_CANDIDATE_CASES))
    for key in ("rustc_verbose", "cargo", "libc", "rustflags"):
        equal(f"performance environment {key}", baseline_env.get(key), candidate_env.get(key))
    for environment in (baseline_env, candidate_env):
        equal("performance contract hash", environment.get("contract_sha256"), digest(HERE / "contract.md"))
    verify_preflight(baseline_dir, CONTROL_CASES)
    verify_preflight(candidate_dir, EXPECTED_CANDIDATE_CASES)
    summary = load(RESULTS / "capture-summary.json")
    equal("performance summary profile before", summary.get("profile_input_sha256_before"), before_inputs)
    equal("performance summary profile after", summary.get("profile_input_sha256_after"), after_inputs)
    equal("performance baseline cleanup", summary.get("cleanup", {}).get("baseline_removed"), True)
    candidate_gate = summary.get("candidate_preflight_gate")
    if not isinstance(candidate_gate, dict):
        raise RuntimeError("performance candidate preflight gate is missing")
    equal("performance candidate preflight status", candidate_gate.get("status"), "ok")
    equal("performance candidate preflight case count", candidate_gate.get("cases"), len(EXPECTED_CANDIDATE_CASES))
    preflight_dir = RESULTS / "preflight-before-timing"
    verify_preflight(preflight_dir, EXPECTED_CANDIDATE_CASES)
    equal("performance candidate gate preflight", candidate_gate.get("preflight"), load(preflight_dir / "preflight.json"))
    equal("performance candidate gate binary", candidate_gate.get("binary_sha256"), candidate_manifest.get("binary_sha256"))
    verify_report(baseline_groups, candidate_groups)
    return {
        "baseline_rows": len(baseline_rows),
        "candidate_rows": len(candidate_rows),
        "baseline_groups": len(baseline_groups),
        "candidate_groups": len(candidate_groups),
        "source_closure": "verified",
        "frozen_lock": GATE_LOCK_SHA256,
    }


def main() -> int:
    receipt = {
        "gates": verify_gates(),
        "oracles": verify_oracle(),
        "native": verify_native(),
        "performance": verify_performance(),
    }
    print(json.dumps(receipt, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
