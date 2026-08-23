use std::{
    fs,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

fn temporary_report_path() -> (std::path::PathBuf, std::path::PathBuf) {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("test clock is after the Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "litchi-perf-xlsx-filesystem-test-{}-{stamp}",
        std::process::id()
    ));
    fs::create_dir(&root).expect("create XLSX filesystem test directory");
    let report = root.join("report.json");
    (root, report)
}

#[test]
fn xlsx_file_selectors_run_as_one_sample_fresh_children() {
    let (root, report) = temporary_report_path();
    let output = Command::new(env!("CARGO_BIN_EXE_litchi-perf-baseline"))
        .args([
            "--case",
            "xlsx_file_open,xlsx_file_open_lifecycle",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--filesystem-cache",
            "warm",
            "--json",
            report.to_str().expect("report path is UTF-8"),
        ])
        .output()
        .expect("run performance harness");
    assert!(
        output.status.success(),
        "harness failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report_value: serde_json::Value =
        serde_json::from_slice(&fs::read(&report).expect("read harness report"))
            .expect("parse harness report");
    let results = report_value["results"]
        .as_array()
        .expect("ordinary result array");
    assert_eq!(
        results.len(),
        2,
        "mixed dispatch must execute each selector once"
    );
    assert_eq!(
        results
            .iter()
            .map(|entry| entry["case"].as_str().expect("result case"))
            .collect::<Vec<_>>(),
        ["xlsx_file_open", "xlsx_file_open_lifecycle"]
    );
    let evidence = report_value["filesystem_evidence"]
        .as_array()
        .expect("filesystem evidence array");
    assert_eq!(evidence.len(), 2);
    for sample in evidence.iter().map(|entry| &entry["samples"][0]) {
        assert_eq!(sample["cache_state"], "warm");
        assert_eq!(
            sample["logical_read_counter_scope"],
            "not_applicable_filesystem_xlsx"
        );
        assert_eq!(
            sample["xlsx_source_sha256"],
            "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036"
        );
        assert_eq!(
            sample["xlsx_semantic_sha256"].as_str().map(str::len),
            Some(64)
        );
    }
    for entry in evidence {
        assert_eq!(entry["fresh_child_per_sample"], true);
        assert_eq!(entry["sample_count"], 1);
    }
    fs::remove_dir_all(root).expect("remove XLSX filesystem test directory");
}

#[test]
fn xlsx_file_cold_verified_is_eligible_or_explicitly_ineligible() {
    let (root, report) = temporary_report_path();
    let output = Command::new(env!("CARGO_BIN_EXE_litchi-perf-baseline"))
        .args([
            "--case",
            "xlsx_file_open",
            "--samples",
            "1",
            "--warmup",
            "0",
            "--filesystem-cache",
            "cold-verified",
            "--json",
            report.to_str().expect("report path is UTF-8"),
        ])
        .output()
        .expect("run cold-verified performance harness");
    assert!(
        output.status.success(),
        "cold-verified harness failed: {}",
        String::from_utf8_lossy(&output.stderr)
    );
    let report_value: serde_json::Value =
        serde_json::from_slice(&fs::read(&report).expect("read cold-verified report"))
            .expect("parse cold-verified report");
    let evidence = &report_value["filesystem_evidence"][0];
    assert!(evidence["cold_verified_status"].is_string());
    let status = evidence["cold_verified_status"]
        .as_str()
        .expect("cold-verified status string");
    if status == "eligible" {
        assert_eq!(evidence["samples"].as_array().map(Vec::len), Some(1));
        assert_eq!(evidence["samples"][0]["cache_state"], "cold-verified");
        assert_eq!(
            evidence["samples"][0]["xlsx_source_sha256"],
            evidence["samples"][0]["cold_verified"]["aligned_source_sha256"]
        );
        assert_eq!(
            evidence["samples"][0]["xlsx_source_sha256"]
                .as_str()
                .map(str::len),
            Some(64)
        );
    } else {
        assert!(evidence["samples"].as_array().is_some_and(Vec::is_empty));
    }
    fs::remove_dir_all(root).expect("remove cold-verified test directory");
}
