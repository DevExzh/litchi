#![cfg(feature = "allocator-metrics")]

use std::{
    fs,
    path::{Path, PathBuf},
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

const CASE: &str = "xlsx_source_backed_cell_values_one_edit_save";
const SAMPLE_COUNT: usize = 2;
const WARMUP_COUNT: usize = 1;
const ALLOCATION_SCOPE: &str = "operation_global_system_allocator";

fn temporary_report_directory() -> PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("test clock is after the Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "litchi-perf-xlsx-planning-allocation-test-{}-{stamp}",
        std::process::id()
    ));
    fs::create_dir(&root).expect("create XLSX planning allocation test directory");
    root
}

fn run_report(binary: &str, report: &Path) -> serde_json::Value {
    let output = Command::new(binary)
        .args([
            "--case",
            CASE,
            "--xlsx-cell-crud-shape",
            "medium",
            "--warmup",
            "1",
            "--samples",
            "2",
            "--json",
            report.to_str().expect("report path is UTF-8"),
        ])
        .output()
        .expect("run XLSX planning allocation harness");
    assert!(
        output.status.success(),
        "{binary} harness failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    serde_json::from_slice(&fs::read(report).expect("read XLSX planning allocation report"))
        .expect("parse XLSX planning allocation report")
}

fn result(report: &serde_json::Value) -> &serde_json::Value {
    assert_eq!(
        report["configuration"]["cases"],
        serde_json::json!([CASE]),
        "the report selected an unexpected case"
    );
    assert_eq!(
        report["configuration"]["xlsx_cell_crud_shapes"],
        serde_json::json!(["medium"]),
        "the report selected an unexpected XLSX cell CRUD shape"
    );
    assert_eq!(
        report["configuration"]["warmup_iterations_per_case"],
        WARMUP_COUNT
    );
    assert_eq!(report["configuration"]["samples_per_case"], SAMPLE_COUNT);

    let results = report["results"].as_array().expect("ordinary result array");
    assert_eq!(results.len(), 1, "the focused run must produce one result");
    assert_eq!(results[0]["case"], CASE);
    results.first().expect("focused result exists")
}

fn xlsx_cell_values_source(result: &serde_json::Value) -> &serde_json::Value {
    result["source"]["xlsx_cell_values"]
        .as_object()
        .map(|_| &result["source"]["xlsx_cell_values"])
        .expect("source-backed XLSX cell-values evidence")
}

fn vector<'a>(source: &'a serde_json::Value, name: &str) -> &'a [serde_json::Value] {
    source[name]
        .as_array()
        .unwrap_or_else(|| panic!("{name} must be an array"))
}

fn unavailable_sample(sample: &serde_json::Value) {
    assert_eq!(
        sample,
        &serde_json::json!({
            "status": "unavailable",
            "scope": ALLOCATION_SCOPE,
        })
    );
}

fn metric(sample: &serde_json::Value, name: &str) -> u64 {
    sample[name]
        .as_u64()
        .unwrap_or_else(|| panic!("measured allocation sample is missing {name}"))
}

fn measured_sample(sample: &serde_json::Value, require_positive_activity: bool) {
    assert_eq!(sample["status"], "measured");
    assert_eq!(sample["scope"], ALLOCATION_SCOPE);

    let allocation_calls = metric(sample, "allocation_calls");
    let allocated_bytes = metric(sample, "allocated_bytes");
    let failed_allocation_calls = metric(sample, "failed_allocation_calls");
    let live_before = metric(sample, "live_bytes_before");
    let live_after = metric(sample, "live_bytes_after");
    let peak_before = metric(sample, "peak_live_bytes_before");
    let peak_after = metric(sample, "peak_live_bytes_after");
    let region_peak = metric(sample, "region_peak_live_bytes");

    for name in [
        "deallocation_calls",
        "reallocation_calls",
        "deallocated_bytes",
    ] {
        let _ = metric(sample, name);
    }
    assert_eq!(failed_allocation_calls, 0);
    if require_positive_activity {
        assert!(
            allocation_calls > 0,
            "planning region recorded no allocations"
        );
        assert!(
            allocated_bytes > 0,
            "planning region recorded no allocated bytes"
        );
    }
    assert!(peak_before >= live_before);
    assert!(peak_after >= live_after);
    assert!(peak_after >= peak_before);
    assert!(region_peak >= live_before);
    assert!(region_peak >= live_after);
    assert!(region_peak <= peak_after);
}

fn assert_phase_sum(result: &serde_json::Value, source: &serde_json::Value) {
    let elapsed = result["elapsed_ns"]["samples"]
        .as_array()
        .expect("elapsed samples array");
    assert_eq!(elapsed.len(), SAMPLE_COUNT);
    let sample_order = result["elapsed_ns"]["sample_order"]
        .as_array()
        .expect("elapsed sample order array");
    assert_eq!(sample_order.len(), SAMPLE_COUNT);
    let mut sorted_order = sample_order
        .iter()
        .map(|value| value.as_u64().expect("sample order index") as usize)
        .collect::<Vec<_>>();
    sorted_order.sort_unstable();
    assert_eq!(sorted_order, (0..SAMPLE_COUNT).collect::<Vec<_>>());

    let open = vector(source, "open_ns");
    let plan = vector(source, "plan_ns");
    let commit = vector(source, "commit_ns");
    let publication = vector(source, "publication_ns");
    for values in [&open, &plan, &commit, &publication] {
        assert_eq!(values.len(), SAMPLE_COUNT);
    }
    for (sorted_index, acquisition_index) in sample_order.iter().enumerate() {
        let acquisition_index = acquisition_index
            .as_u64()
            .expect("sample order acquisition index") as usize;
        let phase_sum = [
            &open[acquisition_index],
            &plan[acquisition_index],
            &commit[acquisition_index],
            &publication[acquisition_index],
        ]
        .into_iter()
        .map(|value| value.as_u64().expect("phase duration"))
        .try_fold(0_u64, u64::checked_add)
        .expect("phase durations fit u64");
        assert_eq!(
            phase_sum,
            elapsed[sorted_index].as_u64().expect("elapsed duration"),
            "phase sum differs at acquisition index {acquisition_index}"
        );
    }
}

#[test]
fn xlsx_source_cell_values_planning_allocations_are_scoped_and_aligned() {
    let root = temporary_report_directory();
    let normal_report = root.join("normal.json");
    let allocator_report = root.join("allocator.json");
    let normal = run_report(env!("CARGO_BIN_EXE_litchi-perf-baseline"), &normal_report);
    let allocator = run_report(
        env!("CARGO_BIN_EXE_litchi-perf-baseline-alloc"),
        &allocator_report,
    );
    let normal_result = result(&normal);
    let allocator_result = result(&allocator);
    let normal_source = xlsx_cell_values_source(normal_result);
    let allocator_source = xlsx_cell_values_source(allocator_result);

    assert_eq!(normal["tool"]["binary"], "litchi-perf-baseline");
    assert_eq!(normal["tool"]["instrumentation"], "none");
    assert_eq!(allocator["tool"]["binary"], "litchi-perf-baseline-alloc");
    assert_eq!(
        allocator["tool"]["instrumentation"],
        "system_allocator_operation_scoped"
    );
    for report in [&normal, &allocator] {
        assert_eq!(report["configuration"]["samples_per_case"], SAMPLE_COUNT);
        assert_eq!(
            report["configuration"]["warmup_iterations_per_case"],
            WARMUP_COUNT
        );
    }
    assert_phase_sum(normal_result, normal_source);
    assert_phase_sum(allocator_result, allocator_source);

    for name in [
        "plan_allocation_metrics",
        "commit_allocation_metrics",
        "publication_allocation_metrics",
    ] {
        assert_eq!(vector(normal_source, name).len(), SAMPLE_COUNT);
        assert_eq!(vector(allocator_source, name).len(), SAMPLE_COUNT);
    }
    for sample in vector(normal_source, "plan_allocation_metrics") {
        unavailable_sample(sample);
    }
    for name in [
        "commit_allocation_metrics",
        "publication_allocation_metrics",
    ] {
        for sample in vector(normal_source, name) {
            unavailable_sample(sample);
        }
    }
    for sample in vector(allocator_source, "plan_allocation_metrics") {
        measured_sample(sample, true);
    }
    for name in [
        "commit_allocation_metrics",
        "publication_allocation_metrics",
    ] {
        for sample in vector(allocator_source, name) {
            measured_sample(sample, false);
        }
    }

    assert_eq!(normal_result["corpus"], allocator_result["corpus"]);
    assert_eq!(
        normal_result["output_sha256"],
        allocator_result["output_sha256"]
    );
    assert_eq!(
        normal_source["output_sha256"],
        allocator_source["output_sha256"]
    );
    assert_eq!(
        normal_source["semantic_sha256"],
        allocator_source["semantic_sha256"]
    );

    fs::remove_dir_all(root).expect("remove XLSX planning allocation test directory");
}
