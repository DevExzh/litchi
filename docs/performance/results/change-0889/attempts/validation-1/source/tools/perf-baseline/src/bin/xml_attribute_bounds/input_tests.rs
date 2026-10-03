//! Input-population and CLI fail-closed checks for the differential harness.

use std::path::{Path, PathBuf};

use serde_json::Value;

use super::{
    InputAccounting, PackageCandidate, collect_packages, is_xml_member, root_has_package,
    validate_differential_args,
};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx")
}

#[test]
fn differential_rejects_unknown_and_incomplete_options() {
    assert!(
        validate_differential_args(&[
            "--json".to_owned(),
            "/tmp/result.json".to_owned(),
            "--generated".to_owned(),
            "0".to_owned(),
            fixture_path().display().to_string(),
        ])
        .is_ok()
    );
    assert!(validate_differential_args(&["--mystery".to_owned(), "x".to_owned()]).is_err());
    assert!(validate_differential_args(&["--json".to_owned()]).is_err());
    assert!(
        validate_differential_args(&[
            "--generated".to_owned(),
            "--aliasing".to_owned(),
            "0".to_owned(),
        ])
        .is_err()
    );
    assert!(
        validate_differential_args(&[
            "--generated".to_owned(),
            "0".to_owned(),
            "--generated".to_owned(),
            "1".to_owned(),
        ])
        .is_err()
    );
}

#[test]
fn package_discovery_deduplicates_overlapping_input_paths() {
    let path = fixture_path();
    let mut packages = Vec::<PackageCandidate>::new();
    let mut accounting = InputAccounting::default();
    collect_packages(&path, &mut packages, &mut accounting).expect("fixture stat must work");
    collect_packages(&path, &mut packages, &mut accounting).expect("fixture stat must work");

    assert_eq!(packages.len(), 1);
    assert_eq!(accounting.scanned_file_count, 2);
    assert_eq!(accounting.failures.len(), 1);
    assert_eq!(
        accounting.failures[0].get("stage").and_then(Value::as_str),
        Some("duplicate_package")
    );
    assert_eq!(accounting.value(&[path])["complete"], false);
}

#[test]
fn member_read_failures_are_retained_with_member_identity() {
    let mut accounting = InputAccounting::default();
    accounting.record_failure(
        "member_read",
        Path::new("fixture.docx"),
        Some("word/document.xml"),
        "CRC mismatch",
    );
    let value = accounting.value(&[Path::new("fixture.docx").to_owned()]);
    assert_eq!(value["read_failure_count"], 1);
    assert_eq!(value["failure_counts"]["member_read"], 1);
    assert_eq!(value["failures"][0]["member"], "word/document.xml");
    assert_eq!(value["complete"], false);
}

#[test]
fn package_selection_matches_the_differential_extensions() {
    assert!(is_xml_member("word/document.xml"));
    assert!(is_xml_member("_rels/.rels"));
    assert!(is_xml_member("drawings/vml.VML"));
    assert!(!is_xml_member("word/document.bin"));
    assert!(!is_xml_member("[Content_Types].json"));
}

#[test]
fn explicit_non_package_root_is_detected_as_incomplete() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("docs/GOAL.md");
    let packages = Vec::<PackageCandidate>::new();
    assert!(!root_has_package(&root, &packages));
}
