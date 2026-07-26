//! Regression tests for workbooks that real producers emit but that a
//! strictly-literal reading of ECMA-376 / MS-XLSB would reject outright.
//!
//! Every fixture here is a complete workbook that previously failed to open.
//! The assertions check that the workbook loads *and* that its content is
//! actually reachable, so a fix cannot regress into "opens but drops data".

use litchi_core::sheet::Worksheet as WorksheetTrait;
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsb::XlsbWorkbook;
use litchi_ooxml::xlsx::Workbook;
use litchi_ooxml::xlsx::external_links::ExternalLinkKind;
use std::fs::File;
use std::path::PathBuf;

/// Resolve a repository-relative fixture path.
fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

/// Excel records how an external link was resolved using the Microsoft
/// `xlExternalLinkPath/*` relationship families. A reader that only accepts
/// the base ECMA-376 `externalLinkPath` type rejects the whole workbook.
///
/// Targets stay inert: opening these workbooks never resolves or contacts the
/// referenced external workbook.
#[test]
fn opens_workbooks_using_microsoft_external_link_path_relationships() {
    for (name, expected_sheets) in [
        ("test-data/ooxml/xlsx/external-link-path-missing.xlsx", 1),
        ("test-data/ooxml/xlsx/external-link-path-startup.xlsx", 1),
    ] {
        let workbook = Workbook::open(fixture(name)).unwrap_or_else(|e| panic!("{name}: {e}"));
        assert_eq!(workbook.worksheet_count(), expected_sheets, "{name}");
    }
}

/// External-link caches address the full column space, so multi-letter
/// columns such as `AM` and `FH` must validate. The previous parser consumed
/// only a single leading letter and rejected every reference beyond `Z`.
#[test]
fn opens_workbook_whose_external_cache_uses_multi_letter_columns() {
    let path = fixture("test-data/ooxml/xlsx/external-cache-multi-letter-column.xlsx");
    let workbook = Workbook::open(&path).unwrap();
    assert!(workbook.worksheet_count() > 0);
    // At least one cached reference must carry a two-letter column, which is
    // exactly what the old single-letter scan could not represent.
    let multi_letter = workbook
        .external_links()
        .iter()
        .filter_map(|link| match &link.kind {
            ExternalLinkKind::Workbook(book) => Some(book),
            _ => None,
        })
        .flat_map(|book| book.cached_sheets.iter())
        .flat_map(|sheet| sheet.rows.iter())
        .flat_map(|row| row.cells.iter())
        .filter_map(|cell| cell.reference.as_deref())
        .any(|reference| {
            reference
                .chars()
                .take_while(|c| c.is_ascii_alphabetic())
                .count()
                > 1
        });
    assert!(
        multi_letter,
        "expected a cached cell reference with a multi-letter column"
    );
}

/// `sheet/@state` outside `ST_SheetState` must not fail the workbook.
/// LibreOffice writes `state="show"` (tdf#118668); the sheets and their names
/// are still fully readable.
#[test]
fn opens_workbook_with_unrecognised_sheet_state() {
    let path = fixture("test-data/ooxml/xlsx/sheet-state-show.xlsx");
    let workbook = Workbook::open(&path).unwrap();
    assert_eq!(workbook.worksheet_count(), 2);
    assert_eq!(workbook.worksheet_names(), ["стр1", "стр2"]);
}

/// A defined name repeated within one scope resolves to a single definition
/// instead of failing the workbook.
#[test]
fn opens_workbook_with_duplicate_defined_names() {
    let path = fixture("test-data/ooxml/xlsx/duplicate-defined-names.xlsx");
    let workbook = Workbook::open(&path).unwrap();
    assert!(workbook.worksheet_count() > 0);
}

/// `sst/@count` and `sst/@uniqueCount` are optional hints. This POI fixture
/// declares `count="8876876876876"`, which does not even fit a `u32`. The
/// hint is ignored and the shared strings themselves stay authoritative.
#[test]
fn opens_workbook_whose_shared_string_count_hint_is_unusable() {
    let path = fixture("test-data/ooxml/xlsx/shared-strings-malformed-count.xlsx");
    let workbook = Workbook::open(&path).unwrap();
    assert!(workbook.worksheet_count() > 0);
    let sheet = workbook.worksheet_by_index(0).unwrap();
    // The shared strings must actually be resolvable, not merely counted.
    assert!(
        sheet.dimensions().is_some(),
        "expected the sheet to expose a populated dimension"
    );
    let mut rows = sheet.rows();
    let mut has_text = false;
    while let Some(row) = rows.next() {
        if row.unwrap().iter().any(|value| {
            matches!(value, litchi_core::sheet::CellValue::String(text) if !text.is_empty())
        }) {
            has_text = true;
            break;
        }
    }
    assert!(has_text, "expected shared-string text to resolve");
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

/// Excel writes `CellParsedFormula` records whose token stream is empty
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
