//! Regression tests for workbooks that real producers emit but that a
//! strictly-literal reading of ECMA-376 / MS-XLSB would reject outright.
//!
//! Every fixture here is a complete workbook that previously failed to open.
//! The assertions check that the workbook loads *and* that its content is
//! actually reachable, so a fix cannot regress into "opens but drops data".

use litchi_core::sheet::Worksheet as WorksheetTrait;
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_xlsb::Workbook as XlsbWorkbook;
use std::fs::File;
use std::path::PathBuf;

/// Resolve a repository-relative fixture path.
fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

/// MS-XLSB part streams may open with `BrtACBegin`/`BrtACEnd` future-record
/// wrapper blocks before the record that starts the part. Skipping them lets
/// the pivot-cache stream parse instead of failing on the first record.
#[test]
fn opens_xlsb_whose_pivot_cache_stream_starts_with_a_future_record_block() {
    let path = fixture("test-data/ooxml/xlsb/pivot-cache-ac-prefixed.xlsb");
    let workbook = XlsbWorkbook::new(File::open(&path).unwrap()).unwrap();
    assert!(workbook.worksheet_count() > 0);
    assert!(
        !workbook.pivot_cache_definitions().is_empty(),
        "expected the AC-prefixed pivot cache definition to be parsed"
    );
}

/// Excel writes `ParsedFormula` records whose token stream is empty
/// (`cce == 0`) even though MS-XLSB 2.5.98.4 requires a positive length.
/// Treating that as fatal rejected the entire workbook over one cell.
#[test]
fn opens_xlsb_containing_an_empty_cell_formula_token_stream() {
    let path = fixture("test-data/ooxml/xlsb/bug66682.xlsb");
    let workbook = XlsbWorkbook::new(File::open(&path).unwrap()).unwrap();
    assert!(workbook.worksheet_count() > 0);
    // The sheet must still be walkable, not merely present.
    let sheet = workbook.worksheet(0).unwrap();
    assert!(WorksheetTrait::dimensions(&sheet).is_some());
}
