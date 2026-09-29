"""Validate the retained 0831 packet before the 0832 experiment is frozen.

This is a read-only gate.  It binds both new 0832 readers to the historical
0831 revision and independently checks the 144 comparative reports' cardinality
and schema fields, including the native no-op allocation omission.  It never
launches a workload or build.
"""

from __future__ import annotations

import json
from pathlib import Path

import analyze as main_reader
import audit as independent_reader
import driver as d


LEGACY = d.ROOT / "docs/performance/results/change-0831"
LEGACY_BASE = "3bcaee6f418a78ee62bab76aed0d6d9f1f93d024"
LEGACY_ANALYSIS_SHA256 = "aaffb3db12a945e24202a38f6edc643e459b5641a7b0e66fa76cc888278d5a58"
LEGACY_AUDIT_SHA256 = "888b3fa361d9f3ed414f3a416b76b045ed80be6538337a311d1a8a53177d68a9"
CASES = (
    ("real-edit", "xlsx_real_file_ordinary_save_edit", None, 500),
    ("real-lifecycle", "xlsx_real_file_ordinary_save_lifecycle", None, 500),
    ("one-cell-tiny", "xlsx_one_cell_commit_save", "tiny", 30),
    ("one-cell-medium", "xlsx_one_cell_commit_save", "medium", 30),
    ("one-cell-dense-wide", "xlsx_one_cell_commit_save", "dense-wide", 30),
    ("one-percent-tiny", "xlsx_one_percent_commit_save", "tiny", 30),
    ("one-percent-medium", "xlsx_one_percent_commit_save", "medium", 30),
    ("one-percent-dense-wide", "xlsx_one_percent_commit_save", "dense-wide", 30),
    ("noop-medium", "xlsx_noop_commit_save", "medium", 500),
)
ALLOCATION_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path):
    require(path.is_file() and not path.is_symlink(), f"missing {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def verify_seal() -> dict[str, str]:
    seal = read(LEGACY / "seal.json")
    require(seal.get("base") == LEGACY_BASE, "0831 seal base changed")
    files = seal.get("files")
    require(isinstance(files, dict) and len(files) > 1000, "0831 seal is incomplete")
    for name in (
        "docs/performance/results/change-0831/analysis.json",
        "docs/performance/results/change-0831/audit.json",
        "docs/performance/results/change-0831/inputs.json",
        "docs/performance/results/change-0831/candidate/after-validation.rs",
    ):
        require(name in files, f"0831 seal omits {name}")
        path = d.ROOT / name
        require(path.is_file() and d.sha(path) == files[name],
                f"0831 sealed file changed: {name}")
    require(d.sha(LEGACY / "analysis.json") == LEGACY_ANALYSIS_SHA256,
            "0831 analysis identity changed")
    require(d.sha(LEGACY / "audit.json") == LEGACY_AUDIT_SHA256,
            "0831 audit identity changed")
    return files


def verify_sealed_file(path: Path, files: dict[str, str], label: str) -> None:
    try:
        name = str(path.resolve().relative_to(d.ROOT))
    except ValueError:
        raise AssertionError(f"{label} escaped repository root: {path}")
    require(name in files, f"0831 seal omits {name}")
    require(path.is_file() and not path.is_symlink() and d.sha(path) == files[name],
            f"0831 sealed file changed: {name}")


def verify_report(path: Path, logical: str, case: tuple[str, str, str | None, int],
                  sealed: dict[str, str]) -> None:
    case_id, case_name, shape, samples = case
    verify_sealed_file(path, sealed, f"{logical} report")
    report = read(path)
    label = str(path)
    expected = {"id": case_id, "case": case_name, "shape": shape,
                "native_samples": samples}
    leg = path.name.split("-")[1]
    require(leg in ("before", "after"), f"{label} leg binding is malformed")
    build_path = LEGACY / f"build-{leg}.json"
    verify_sealed_file(build_path, sealed, f"{logical} build-{leg}")
    binary = read(build_path)["binaries"][logical]
    warmup = 3 if logical == "native" else 0
    # Exercise both new 0832 parser implementations against the retained
    # 0831 raw reports.  ``historical_base`` is an explicit preflight-only
    # binding; ordinary packet replay uses the current driver base.
    main_reader.validate_report(path, expected, logical, samples, warmup, binary, label,
                                historical_base=LEGACY_BASE)
    independent_reader.validate_report(path, expected, logical, samples, warmup, binary, label,
                                       expected_revision=LEGACY_BASE)
    require(report.get("schema_version") == 1, f"{label} schema changed")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline",
            f"{label} tool identity changed")
    expected_tool = "litchi-perf-baseline" if logical == "native" else "litchi-perf-baseline-alloc"
    require(tool.get("binary") == expected_tool, f"{label} binary tool changed")
    if logical == "native":
        require(tool.get("instrumentation") == "none", f"{label} instrumentation changed")
    else:
        require(tool.get("instrumentation")
                == "ordinary_save_procfs_and_system_allocator_operation_scoped",
                f"{label} instrumentation identity changed")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case")
            == (3 if logical == "native" else 0)
            and configuration.get("cases") == [case_name],
            f"{label} configuration changed")
    if shape is not None:
        shapes = configuration.get("xlsx_shapes")
        require(isinstance(shapes, list) and shape in shapes, f"{label} shape changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label} result count changed")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == case_name,
            f"{label} result case changed")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label} elapsed schema changed")
    values = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(values, list) and len(values) == samples and values == sorted(values)
            and all(type(value) is int and value > 0 for value in values),
            f"{label} elapsed vector changed")
    require(isinstance(order, list) and len(order) == samples
            and sorted(order) == list(range(samples)), f"{label} sample order changed")
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict) and metrics.get("sample_count") == samples
            and metrics.get("sample_indices") == order,
            f"{label} operation metrics changed")
    allocation = metrics.get("allocation")
    if logical == "native" and case_id == "noop-medium":
        require("allocation" not in metrics, f"{label} native no-op allocation appeared")
    else:
        require(isinstance(allocation, dict), f"{label} allocation envelope missing")
        if logical == "observer":
            require(allocation.get("status") == "measured", f"{label} allocation status changed")
            for field in ALLOCATION_FIELDS:
                metric = allocation.get(field)
                require(isinstance(metric, dict) and metric.get("status") == "measured"
                        and isinstance(metric.get("values"), list)
                        and len(metric["values"]) == samples,
                        f"{label} allocation vector changed: {field}")
    rss = path.with_suffix(".rss")
    verify_sealed_file(rss, sealed, f"{label} RSS")
    require(rss.is_file() and not rss.is_symlink() and rss.read_text().strip().isdigit()
            and int(rss.read_text().strip()) > 0, f"{label} RSS sidecar changed")


def verify_reports(sealed: dict[str, str]) -> int:
    count = 0
    for logical, blocks, samples in (("native", 6, None), ("observer", 2, 3)):
        for block in range(blocks):
            for leg in ("before", "after"):
                for case in CASES:
                    expected_samples = case[3] if logical == "native" else samples
                    path = LEGACY / logical / f"{block:02d}-{leg}-{case[0]}.json"
                    verify_report(path, logical, (case[0], case[1], case[2], expected_samples), sealed)
                    count += 1
    require(count == 144, f"0831 comparative report count changed: {count}")
    return count


def main() -> None:
    sealed = verify_seal()
    reports = verify_reports(sealed)
    print(json.dumps({"status": "pass", "historical_packet": "0831",
                      "comparative_reports": reports,
                      "readers": ["0832-analyze", "0832-audit"]},
                     sort_keys=True))


if __name__ == "__main__":
    main()
