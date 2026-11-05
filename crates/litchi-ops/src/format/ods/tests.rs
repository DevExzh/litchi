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

/// The byte route's first arm: the ODF package an ordinary ODS prepares,
/// which the byte opener adopts (`prepared_or_original`).
fn open_prepared(bytes: Vec<u8>) -> Ods {
    let prepared = litchi_odf_common::detect::prepared_or_original(
        bytes,
        litchi_odf_common::SourceLimits::default(),
    )
    .unwrap_or_else(|_| panic!("expected a prepared ODS package"));
    Ods::from_prepared(prepared, &OpenOptions::default()).expect("open prepared ODS")
}

/// The facade's `from_detected` arm, which re-parsed the bytes detection
/// retained: the unified byte route adopts the prepared package instead.
fn open_detected(bytes: Vec<u8>) -> Ods {
    assert!(matches!(
        detect_format_smart(bytes.clone(), &crate::detect::SourceLimits::default()),
        Some(DetectedFormat::Ods(_))
    ));
    match crate::open_bytes(bytes).expect("open detected ODS") {
        crate::File::Ods(handle) => handle,
        other => panic!("expected ODS, got {other:?}"),
    }
}

/// The path route: [`crate::open`] builds the source-backed owner.
#[cfg(any(unix, windows))]
fn open_source(path: &Path) -> Ods {
    match crate::open(path).expect("open ODS path") {
        crate::File::Ods(handle) => handle,
        other => panic!("expected ODS, got {other:?}"),
    }
}

fn assert_matches_leaf(ods: &Ods, leaf: &litchi_ods::Workbook, name: &str) {
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
    assert!(Arc::ptr_eq(&ods.0.inner, &clone.0.inner));
    assert_eq!(format!("{ods:?}"), r#"Ods { owner: "eager", .. }"#);
}

#[test]
fn byte_routes_match_the_leaf_reader() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).expect("read ODS fixture");
        let leaf = Package::from_bytes(&bytes)
            .expect("leaf ODS")
            .workbook()
            .expect("leaf ODS");
        assert!(leaf.sheet_count() > 0, "{name}");
        for ods in [open_prepared(bytes.clone()), open_detected(bytes)] {
            assert!(matches!(ods.owner(), OdsOwner::Eager { .. }));
            assert_matches_leaf(&ods, &leaf, name);
        }
    }
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_matches_the_leaf_reader() {
    for name in FIXTURES {
        let path = fixture(name);
        let leaf = Package::from_bytes(&std::fs::read(&path).expect("read"))
            .expect("leaf ODS")
            .workbook()
            .expect("leaf ODS");
        let ods = open_source(&path);
        assert!(matches!(ods.owner(), OdsOwner::Source { .. }));
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
fn malformed_body_is_a_typed_odf_error() {
    let bytes = broken_package(litchi_odf_common::constants::ODF_SPREADSHEET);
    let Err(leaf_error) = Package::from_bytes(&bytes).and_then(|package| package.workbook()) else {
        panic!("a broken body must not open in the leaf");
    };
    let error = crate::open_bytes(bytes).expect_err("a broken body must not open");
    assert!(
        matches!(error.leaf(), crate::error::Source::Core(_)),
        "{error:?}"
    );
    assert_eq!(error.format(), Some(FileFormat::Ods));
    assert_eq!(error.message(), leaf_error.to_string());
}

/// A package with a valid `mimetype` and manifest and an unparsable body.
fn broken_package(mimetype: &str) -> Vec<u8> {
    use std::io::Write as _;

    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("mimetype", options).unwrap();
    writer.write_all(mimetype.as_bytes()).unwrap();
    writer.start_file("META-INF/manifest.xml", options).unwrap();
    writer
        .write_all(
            format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="{mimetype}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer.start_file("content.xml", options).unwrap();
    writer.write_all(b"<office:document-content>").unwrap();
    writer.finish().unwrap().into_inner()
}

/// Record 0287: a walk that wants only cells lists a sheet's content cells
/// directly; it must find exactly the cells, in the same order and with the
/// same paths and previews, that the phased sheet walk finds when it also
/// lists rows, columns and merges.
#[test]
fn cell_only_walks_match_the_phased_sheet_walk() {
    let mut fixtures: Vec<&str> = FIXTURES.to_vec();
    fixtures.extend([
        "odf/corpus/calc-cell-styles.ods",
        "odf/ods/grid-padding-package.ods",
        "odfdo/tests/samples/simple_table_named_range.ods",
    ]);
    for name in fixtures {
        let file =
            crate::open_bytes(std::fs::read(fixture(name)).expect("fixture")).expect("open ODS");
        let describe = |node: crate::node::Node<'_>| {
            let view = node.view();
            (
                node.kind(),
                node.path().to_string(),
                view.preview().expect("preview").map(|p| p.to_string()),
                view.text().expect("text").map(|text| text.into_owned()),
            )
        };
        let direct: Vec<_> = file
            .query("cell")
            .expect("query")
            .map(|node| describe(node.expect("node")))
            .collect();
        let phased: Vec<_> = file
            .query("cell, row, col, mergeCell")
            .expect("query")
            .map(|node| describe(node.expect("node")))
            .filter(|(kind, ..)| *kind == "cell")
            .collect();
        assert!(!direct.is_empty(), "{name}: no cells");
        assert_eq!(direct, phased, "{name}");
    }
}

/// Record 0287: the binary-searched row runs find exactly the run
/// `Sheet::cell` finds, at every run boundary of every fixture sheet
/// (including the huge padding runs producers write).
#[test]
fn row_ends_find_the_runs_sheet_cell_finds() {
    let mut fixtures: Vec<&str> = FIXTURES.to_vec();
    fixtures.extend([
        "odf/corpus/calc-cell-styles.ods",
        "odf/ods/grid-padding-package.ods",
        "odfdo/tests/samples/simple_table_named_range.ods",
    ]);
    for name in fixtures {
        let file =
            crate::open_bytes(std::fs::read(fixture(name)).expect("fixture")).expect("open ODS");
        let crate::File::Ods(ods) = &file else {
            panic!("{name}: expected ODS");
        };
        for (position, sheet) in ods.worksheets().expect("sheets").iter().enumerate() {
            let ends = node::row_ends(ods, position as u32).expect("row ends");
            let mut rows = vec![0_usize];
            let mut end = 0_usize;
            for row in &sheet.rows {
                rows.extend([end.saturating_sub(1), end, end.saturating_add(1)]);
                end = end.saturating_add(row.repeat());
            }
            rows.extend([
                end.saturating_sub(1),
                end,
                end.saturating_add(7),
                usize::MAX,
            ]);
            let mut columns = vec![0_usize, 1, 2, 3, 1023, 1024, 16_383, 16_384];
            if let Some(row) = sheet.rows.first() {
                let mut end = 0_usize;
                for cell in &row.cells {
                    columns.extend([end.saturating_sub(1), end, end.saturating_add(1)]);
                    end = end.saturating_add(cell.repeat());
                }
            }
            for &row in &rows {
                for &column in &columns {
                    let expected = sheet.cell(row, column);
                    let found = ends.cell(sheet, row, column);
                    assert!(
                        match (expected, found) {
                            (Some(expected), Some(found)) => std::ptr::eq(expected, found),
                            (None, None) => true,
                            _ => false,
                        },
                        "{name} sheet {position}: ({row}, {column})"
                    );
                }
            }
        }
    }
}
