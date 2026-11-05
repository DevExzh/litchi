//! The paired unified/leaf XLSX routes of M0a step S8b run through the same
//! fresh-child filesystem protocol as their predecessors, report the same
//! source and semantic identities, and say which pair they belong to.

use std::{
    fs,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

const XLSX_SOURCE_SHA256: &str = "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036";
const XLSX_SEMANTIC_SHA256: &str =
    "020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e";
const SELECTED_CELL_EVIDENCE_DIGEST: &str =
    "36e53d9002ae8c433ad918b400196fb886fa675f850076808ac51327d1f42ac1";

#[test]
fn paired_xlsx_routes_run_in_fresh_children_and_name_their_pairs() {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("test clock is after the Unix epoch")
        .as_nanos();
    let root = std::env::temp_dir().join(format!(
        "litchi-perf-unified-ops-test-{}-{stamp}",
        std::process::id()
    ));
    fs::create_dir(&root).expect("create test directory");
    let report = root.join("report.json");
    let filesystem_cases = [
        ("xlsx_file_unified_open", "xlsx_file_open", "unified"),
        ("xlsx_file_leaf_open", "xlsx_file_open", "leaf"),
        (
            "xlsx_file_unified_open_lifecycle",
            "xlsx_file_open_lifecycle",
            "unified",
        ),
        (
            "xlsx_file_leaf_open_lifecycle",
            "xlsx_file_open_lifecycle",
            "leaf",
        ),
        (
            "xlsx_file_unified_selected_cell",
            "xlsx_file_selected_cell",
            "unified",
        ),
        (
            "xlsx_file_leaf_selected_cell",
            "xlsx_file_selected_cell",
            "leaf",
        ),
    ];
    let in_process_cases = [
        ("xlsx_bytes_unified_open", "xlsx_bytes_open", "unified"),
        ("xlsx_bytes_leaf_open", "xlsx_bytes_open", "leaf"),
    ];
    let selection = filesystem_cases
        .iter()
        .chain(&in_process_cases)
        .map(|(name, _, _)| *name)
        .collect::<Vec<_>>()
        .join(",");
    let output = Command::new(env!("CARGO_BIN_EXE_litchi-perf-baseline"))
        .args([
            "--case",
            &selection,
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
    fs::remove_dir_all(&root).expect("remove test directory");

    // Filesystem pairs report first, in selection order, then the in-process
    // pairs.
    let results = report_value["results"].as_array().expect("result array");
    assert_eq!(
        results.len(),
        filesystem_cases.len() + in_process_cases.len()
    );
    for (result, (name, predecessor, leg)) in results
        .iter()
        .zip(filesystem_cases.iter().chain(&in_process_cases))
    {
        assert_eq!(result["case"], *name);
        let pair = &result["source"]["unified_ops"];
        assert_eq!(pair["leg"], *leg, "{name}");
        assert_eq!(pair["predecessor"], *predecessor, "{name}");
        let twin = match *leg {
            "unified" => name.replacen("_unified_", "_leaf_", 1),
            _ => name.replacen("_leaf_", "_unified_", 1),
        };
        assert_eq!(pair["twin"], twin.as_str(), "{name}");
        assert!(
            pair["timed_calls"]
                .as_str()
                .is_some_and(|calls| !calls.is_empty())
        );
    }
    for result in &results[..filesystem_cases.len()] {
        assert_eq!(result["cache_state"], "warm");
    }

    let evidence = report_value["filesystem_evidence"]
        .as_array()
        .expect("filesystem evidence array");
    assert_eq!(evidence.len(), filesystem_cases.len());
    for (entry, (name, _, _)) in evidence.iter().zip(&filesystem_cases) {
        assert_eq!(entry["case"], *name);
        assert_eq!(entry["fresh_child_per_sample"], true);
        let samples = entry["samples"].as_array().expect("sample array");
        assert_eq!(samples.len(), 1);
        let sample = &samples[0];
        assert_eq!(
            sample["logical_read_counter_scope"],
            "not_applicable_filesystem_xlsx"
        );
        assert_eq!(sample["xlsx_source_sha256"], XLSX_SOURCE_SHA256);
        assert_eq!(sample["xlsx_semantic_sha256"], XLSX_SEMANTIC_SHA256);
        if name.ends_with("_selected_cell") {
            assert_eq!(
                sample["xlsx_selected_cell"]["digest"],
                SELECTED_CELL_EVIDENCE_DIGEST
            );
        } else {
            assert!(sample.get("xlsx_selected_cell").is_none());
        }
    }
}
