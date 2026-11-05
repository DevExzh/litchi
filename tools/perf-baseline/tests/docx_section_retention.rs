#![cfg(feature = "allocator-metrics")]

use std::fs;
use std::path::PathBuf;
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

const SOURCE_REVISION: &str = "0123456789abcdef0123456789abcdef01234567";

fn temporary_report() -> (PathBuf, PathBuf) {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("test clock is after the Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "litchi-perf-docx-section-retention-test-{}-{stamp}",
        std::process::id()
    ));
    fs::create_dir(&root).expect("create section-retention test directory");
    let report = root.join("report.json");
    (root, report)
}

fn run_report(
    sections: usize,
    concurrent_commits: usize,
    policy: &str,
    ownership: &str,
) -> (PathBuf, serde_json::Value) {
    let (root, report) = temporary_report();
    let output = Command::new(env!("CARGO_BIN_EXE_litchi-perf-baseline-alloc"))
        .args([
            "docx-section-retention",
            "--sections",
            &sections.to_string(),
            "--concurrent-commits",
            &concurrent_commits.to_string(),
            "--policy",
            policy,
            "--ownership",
            ownership,
            "--source-revision",
            SOURCE_REVISION,
            "--output",
            report.to_str().expect("report path is UTF-8"),
        ])
        .output()
        .expect("run DOCX section-retention diagnostic");
    assert!(
        output.status.success(),
        "section-retention diagnostic failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let value = serde_json::from_slice(&fs::read(&report).expect("read section-retention report"))
        .expect("parse section-retention report");
    (root, value)
}

fn u64_field(value: &serde_json::Value, field: &str) -> u64 {
    value[field]
        .as_u64()
        .unwrap_or_else(|| panic!("{field} must be an unsigned integer"))
}

fn assert_gauge_sequence(scenario: &serde_json::Value) {
    let gauges = &scenario["gauges"];
    let phases = [
        &gauges["baseline"],
        &gauges["after_snapshot_drop"]["absolute"],
        &gauges["commits_retained"]["absolute"],
        &gauges["after_release_all_but_last"]["absolute"],
        &gauges["after_release_last"]["absolute"],
        &gauges["after_commit_drop"]["absolute"],
    ];
    for phase in phases {
        assert_eq!(phase["overflowed"], false);
        assert_eq!(phase["observer_invalid"], false);
        assert!(
            u64_field(phase, "peak_live_bytes") >= u64_field(phase, "live_bytes"),
            "peak live bytes must cover current live bytes"
        );
    }
    for pair in phases.windows(2) {
        for field in [
            "allocation_calls",
            "deallocation_calls",
            "reallocation_calls",
            "failed_allocation_calls",
            "allocated_bytes",
            "deallocated_bytes",
            "peak_live_bytes",
        ] {
            assert!(
                u64_field(pair[1], field) >= u64_field(pair[0], field),
                "allocator counter {field} regressed"
            );
        }
        let live = u64_field(pair[0], "live_bytes")
            .checked_add(
                u64_field(pair[1], "allocated_bytes")
                    .checked_sub(u64_field(pair[0], "allocated_bytes"))
                    .expect("allocated bytes regressed"),
            )
            .and_then(|value| {
                value.checked_sub(
                    u64_field(pair[1], "deallocated_bytes")
                        .checked_sub(u64_field(pair[0], "deallocated_bytes"))
                        .expect("deallocated bytes regressed"),
                )
            })
            .expect("live-byte transition overflowed");
        assert_eq!(live, u64_field(pair[1], "live_bytes"));
    }
    assert!(
        gauges["commits_retained"]["absolute"]["region_peak_live_bytes"].is_null(),
        "the unjournaled correctness checkpoint must not claim a region peak"
    );
    for phase_name in [
        "after_snapshot_drop",
        "after_release_all_but_last",
        "after_release_last",
        "after_commit_drop",
    ] {
        let phase = &gauges[phase_name];
        let region_peak = u64_field(&phase["absolute"], "region_peak_live_bytes");
        assert!(
            region_peak >= u64_field(&phase["absolute"], "live_bytes"),
            "region peak must cover phase endpoint for {phase_name}"
        );
    }
    assert_eq!(gauges["returned_to_baseline"], true);
    assert_eq!(
        u64_field(&gauges["after_commit_drop"]["absolute"], "live_bytes"),
        u64_field(&gauges["baseline"], "live_bytes")
    );
}

fn assert_report_shape(
    report: &serde_json::Value,
    sections: usize,
    concurrent_commits: usize,
    policy: &str,
    ownership: &str,
) {
    assert_eq!(report["schema"], "docx_section_retention_v1");
    assert_eq!(report["section_count"], sections);
    assert_eq!(report["concurrent_commits"], concurrent_commits);
    assert_eq!(report["retention_policy"], policy);
    assert_eq!(report["ownership_mode"], ownership);
    assert_eq!(report["source_revision"], SOURCE_REVISION);
    assert!(
        report["source_archive_bytes"]
            .as_u64()
            .is_some_and(|v| v > 0)
    );
    assert_eq!(
        report["source_archive_sha256"].as_str().map(str::len),
        Some(64)
    );
    assert_eq!(
        report["expected_output_sha256"]
            .as_array()
            .expect("expected output digests")
            .len(),
        concurrent_commits
    );
    assert_eq!(
        report["expected_output_bytes"]
            .as_array()
            .expect("expected output lengths")
            .len(),
        concurrent_commits
    );

    let admission = &report["retention_admission"];
    assert_eq!(admission["boundary_checks_verified"], false);
    assert_eq!(u64_field(admission, "required_retained_base_bytes"), 0);
    assert_eq!(u64_field(admission, "zero_retained_base_bytes"), 0);
    assert_eq!(u64_field(admission, "one_under_retained_base_bytes"), 0);
    assert_eq!(u64_field(admission, "exact_retained_base_bytes"), 0);

    let scenarios = report["scenarios"].as_array().expect("scenario array");
    assert_eq!(scenarios.len(), 1);
    let scenario = &scenarios[0];
    assert_eq!(scenario["ownership_mode"], ownership);
    assert_eq!(scenario["section_count"], sections);
    assert_eq!(scenario["concurrent_commits"], concurrent_commits);
    assert_eq!(scenario["source_inventory_sharing_verified"], true);
    let pair_count = concurrent_commits * concurrent_commits.saturating_sub(1) / 2;
    assert_eq!(
        u64_field(scenario, "source_inventory_pair_count"),
        pair_count as u64
    );
    assert_eq!(
        u64_field(scenario, "source_inventory_shared_pairs"),
        if ownership == "shared_snapshot_clones" {
            pair_count as u64
        } else {
            0
        }
    );
    assert_eq!(scenario["measured_commit_shape_verified"], true);
    assert_eq!(scenario["measured_output_verified"], true);
    assert_eq!(scenario["measured_inverse_verified"], true);
    assert_eq!(scenario["preflight_output_verified"], true);
    assert_eq!(scenario["preflight_inverse_verified"], true);
    assert_eq!(
        u64_field(scenario, "measured_commit_validation_count"),
        if ownership == "shared_snapshot_clones" {
            1
        } else {
            concurrent_commits as u64
        }
    );
    assert_gauge_sequence(scenario);
}

#[test]
fn small_shared_exact_retention_preserves_oracles_and_sharing() {
    let (root, report) = run_report(2, 2, "exact", "shared_snapshot_clones");
    assert_report_shape(&report, 2, 2, "exact", "shared_snapshot_clones");
    let admission = &report["retention_admission"];
    let scenario = &report["scenarios"][0];
    assert_eq!(
        scenario["retention_ceiling_bytes"],
        admission["exact_cap_bytes"]
    );
    let retained = scenario["retained_base_bytes_each"]
        .as_array()
        .expect("retained-byte vector");
    assert_eq!(retained.len(), 2);
    for value in retained {
        assert_eq!(value, &serde_json::json!(0));
    }
    assert_eq!(scenario["retained_commit_count"], 0);
    fs::remove_dir_all(root).expect("remove small section-retention test directory");
}

#[test]
fn large_independent_one_under_cap_records_fallback_without_changing_oracles() {
    let (root, report) = run_report(8_192, 1, "one_under", "independent_sources");
    assert_report_shape(&report, 8_192, 1, "one_under", "independent_sources");
    let admission = &report["retention_admission"];
    let scenario = &report["scenarios"][0];
    assert_eq!(
        scenario["retention_ceiling_bytes"],
        admission["one_under_cap_bytes"]
    );
    assert_eq!(
        scenario["retained_base_bytes_each"],
        serde_json::json!([0]),
        "one-under admission must fall back without retained state"
    );
    assert_eq!(scenario["retained_commit_count"], 0);
    assert_eq!(scenario["logical_retained_base_bytes_sum"], 0);
    fs::remove_dir_all(root).expect("remove large section-retention test directory");
}
