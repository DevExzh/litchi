use super::*;
use crate::Code;
use crate::detect::{DetectedFormat, detect_format_smart};
use std::path::{Path, PathBuf};

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

const FIXTURES: [&str; 3] = [
    "odf/corpus/calc-two-sheets.ods",
    "odf/corpus/calc-formulas.ods",
    "odf/corpus/calc-unicode-chinese.ods",
];

/// The facade's first byte-route arm: checked ODS preparation.
fn open_prepared(bytes: Vec<u8>) -> Ods {
    let prepared = crate::detect::detected::detect_prepared_ods(bytes)
        .unwrap_or_else(|_| panic!("expected a prepared ODS package"));
    Ods::from_prepared(prepared, &OpenOptions::default()).expect("open prepared ODS")
}

/// The facade's `from_detected` arm: smart detection retains the bytes.
fn open_detected(bytes: Vec<u8>) -> Ods {
    match detect_format_smart(bytes) {
        Some(DetectedFormat::Ods(bytes)) => {
            Ods::from_bytes(bytes, &OpenOptions::default()).expect("open detected ODS")
        },
        other => panic!("expected ODS detection, got {other:?}"),
    }
}

/// The facade's path route: the workbook path probe builds the source owner.
#[cfg(any(unix, windows))]
fn open_source(path: &Path) -> Ods {
    match crate::detect::detected::detect_workbook_source_path(path).expect("probe ODS path") {
        crate::detect::detected::WorkbookSourcePathDetection::Ods(spreadsheet) => {
            Ods::from_source(spreadsheet, &OpenOptions::default()).expect("open ODS source")
        },
        _ => panic!("expected an ODS source owner"),
    }
}

fn assert_matches_leaf(ods: &Ods, leaf: &litchi_ods::Spreadsheet, name: &str) {
    assert_eq!(
        ods.worksheet_names().expect("names"),
        leaf.sheet_names(),
        "{name}"
    );
    assert_eq!(
        ods.worksheet_count().expect("count"),
        leaf.sheet_count(),
        "{name}"
    );
    assert_eq!(
        ods.worksheets().expect("sheets").len(),
        leaf.sheets().len(),
        "{name}"
    );
    assert_eq!(
        ods.text().expect("text"),
        leaf.text().expect("leaf text"),
        "{name}"
    );
    assert_eq!(
        format!("{:?}", ods.metadata().expect("metadata").expect("Some")),
        format!("{:?}", leaf.metadata()),
        "{name}"
    );
    assert_eq!(
        &ods.tolerance_report().expect("report"),
        leaf.tolerance_report(),
        "{name}"
    );
}

#[test]
fn handle_is_one_pointer_and_shares_its_state() {
    assert_eq!(size_of::<Ods>(), size_of::<usize>());
    let ods = open_prepared(std::fs::read(fixture(FIXTURES[0])).expect("read"));
    let clone = ods.clone();
    assert!(Arc::ptr_eq(&ods.0, &clone.0));
    assert_eq!(format!("{ods:?}"), r#"Ods { owner: "eager", .. }"#);
}

#[test]
fn byte_routes_match_the_leaf_reader() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).expect("read ODS fixture");
        let leaf = litchi_ods::Spreadsheet::from_bytes(bytes.clone()).expect("leaf ODS");
        assert!(leaf.sheet_count() > 0, "{name}");
        for ods in [open_prepared(bytes.clone()), open_detected(bytes)] {
            assert!(matches!(ods.owner(), OdsOwner::Eager(_)));
            assert_matches_leaf(&ods, &leaf, name);
        }
    }
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_matches_the_leaf_reader() {
    for name in FIXTURES {
        let path = fixture(name);
        let leaf = litchi_ods::Spreadsheet::from_bytes(std::fs::read(&path).expect("read"))
            .expect("leaf ODS");
        let ods = open_source(&path);
        assert!(matches!(ods.owner(), OdsOwner::Source(_)));
        assert_eq!(format!("{ods:?}"), r#"Ods { owner: "source", .. }"#);
        assert_matches_leaf(&ods, &leaf, name);
    }
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_reports_a_changed_source_as_io_error() {
    let temporary = tempfile::Builder::new()
        .suffix(".ods")
        .tempfile()
        .expect("temporary ODS");
    let bytes = std::fs::read(fixture(FIXTURES[0])).expect("read");
    std::fs::write(temporary.path(), &bytes).expect("write ODS");
    let ods = open_source(temporary.path());

    let mut changed = bytes;
    changed.extend_from_slice(b"trailing bytes change the source revision");
    std::fs::write(temporary.path(), changed).expect("rewrite ODS");

    for error in [
        ods.worksheet_names().expect_err("names after change"),
        ods.worksheet_count().expect_err("count after change"),
        ods.worksheets().expect_err("sheets after change"),
        ods.text().expect_err("text after change"),
        ods.metadata().expect_err("metadata after change"),
    ] {
        assert_eq!(error.code(), Code::IoError, "{error}");
        assert!(matches!(error.leaf(), crate::error::Source::Core(_)));
        assert_eq!(error.format(), Some(FileFormat::Ods));
    }
}

#[test]
fn malformed_body_is_a_typed_core_error() {
    let Err(leaf_error) = litchi_ods::Spreadsheet::from_bytes(b"PK\x03\x04 not a zip".to_vec())
    else {
        panic!("broken package must not open in the leaf");
    };
    let error = Ods::from_bytes(b"PK\x03\x04 not a zip".to_vec(), &OpenOptions::default())
        .expect_err("broken package must not open");
    assert!(
        matches!(error.leaf(), crate::error::Source::Core(_)),
        "{error:?}"
    );
    assert_eq!(error.format(), Some(FileFormat::Ods));
    assert_eq!(error.message(), leaf_error.to_string());
}
