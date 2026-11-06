//! Native inert macro-sheet catalogs and edits to the ordinary worksheet.
use litchi_biff::{Limits, Records};
use litchi_core::sheet::{Cell as _, CellValue, WorkbookTrait};
use litchi_xls::{Package, SheetKind, SheetVisibility, SourceBackedWorkbook, raw::Reader};
use std::io::Cursor;
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel/excel-legacy-sheets.xls"),
    )
    .unwrap()
}
fn workbook(bytes: &[u8]) -> Vec<u8> {
    litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap()
        .open_stream(&["Workbook"])
        .unwrap()
}
fn macro_stream(bytes: &[u8]) -> (usize, &[u8]) {
    let start = Records::new(bytes, Limits::default())
        .unwrap()
        .map(|r| r.unwrap())
        .find(|r| r.kind().get() == 0x85 && r.payload()[5] == 1)
        .map(|r| u32::from_le_bytes(r.payload()[..4].try_into().unwrap()) as usize)
        .unwrap();
    let end = Records::new(&bytes[start..], Limits::default())
        .unwrap()
        .map(|r| r.unwrap())
        .find(|r| r.kind().get() == 0x0a)
        .map(|r| r.offset() + 4)
        .unwrap();
    (start, &bytes[start..start + end])
}
fn check_catalog(bytes: &[u8]) {
    let eager = Reader::new(Cursor::new(bytes)).unwrap();
    let source = SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
        litchi_core::OwnedSource::new(bytes.to_vec()),
    ))
    .unwrap();
    let metadata = source.sheet_metadata().unwrap();
    assert_eq!(metadata, eager.sheets());
    assert_eq!(metadata.len(), 2);
    for (i, s) in metadata.iter().enumerate() {
        assert_eq!(s.name(), ["Macro_016", "Sheet1"][i]);
        assert_eq!(s.workbook_index(), i);
        assert_eq!(s.visibility(), SheetVisibility::Visible);
    }
    assert_eq!(metadata[0].kind(), SheetKind::MacroSheet);
    assert_eq!(metadata[0].parsed_worksheet_index(), None);
    assert_eq!(metadata[1].kind(), SheetKind::WorksheetOrDialog);
    assert_eq!(metadata[1].parsed_worksheet_index(), Some(0));
    assert_eq!(source.sheet("Macro_016").unwrap().unwrap(), metadata[0]);
    assert_eq!(source.worksheet_names().unwrap(), ["Sheet1"]);
    assert_eq!(eager.worksheet_names(), ["Sheet1"]);
    assert!(source.worksheet_by_name("Macro_016").unwrap().is_none());
    assert!(
        matches!(eager.worksheet_by_name("Macro_016"),Err(litchi_xls::Error::WorksheetNotFound(name)) if name=="Macro_016")
    );
    let sheet = eager.worksheet_by_name("Sheet1").unwrap();
    // The opaque macro cells reference these shared strings. Ordinary-sheet
    // publication must not drop or reuse their slots.
    assert_eq!(
        &sheet.shared_strings().unwrap()[4..7],
        [
            "MACROSHEET_MARKER_016",
            "INERT_LABEL_ONLY_016",
            "END_MACROSHEET_016"
        ]
    );
    for (r, c, text) in [
        (0, 0, "EXCEL_LEGACY_SHEETS_016"),
        (0, 1, "Fixture ID: excel-legacy-sheets-016"),
        (1, 0, "INERT_CONTENT_016"),
        (1, 1, "No macro formulas or code are executed."),
    ] {
        assert_eq!(
            sheet.get_cell(r, c).unwrap().value(),
            &CellValue::String(text.into())
        );
    }
}
#[test]
fn native_macro_catalog_retains_a_real_nonworksheet_substream() {
    let bytes = native();
    check_catalog(&bytes);
    let wb = workbook(&bytes);
    let (start, data) = macro_stream(&wb);
    assert_eq!(start, 15466);
    assert_eq!(data.len(), 510);
    let records = Records::new(data, Limits::default())
        .unwrap()
        .map(|r| r.unwrap())
        .collect::<Vec<_>>();
    assert_eq!(records.len(), 38);
    assert_eq!(
        u16::from_le_bytes(records[0].payload()[2..4].try_into().unwrap()),
        0x0040
    );
    let cells = records
        .iter()
        .filter(|r| r.kind().get() == 0x00fd)
        .map(|r| {
            let b = r.payload();
            (
                u16::from_le_bytes(b[..2].try_into().unwrap()),
                u16::from_le_bytes(b[2..4].try_into().unwrap()),
                u32::from_le_bytes(b[6..10].try_into().unwrap()),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(cells, [(0, 0, 4), (1, 0, 5), (2, 0, 6)]);
    assert!(!records.iter().any(|r| r.kind().get() == 0x0006));
    assert!(Package::from_bytes(&bytes).unwrap().to_bytes().unwrap() == bytes);
}
#[test]
fn native_ordinary_sheet_text_edit_preserves_macro_records_and_index_targets() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    edit.insert_cell(
        "Sheet1",
        litchi_xls::cell_values::Reference::new(3, 3).unwrap(),
        litchi_xls::cell_values::Value::Text("Reviewed ordinary worksheet".into()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let mut old =
        litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    let mut new =
        litchi_cfb::OleFile::open(Cursor::new(&after), litchi_cfb::OleFileLimits::default())
            .unwrap();
    assert_eq!(old.list_streams(), new.list_streams());
    assert!(!old.exists(&["_VBA_PROJECT_CUR"]));
    for path in old.list_streams() {
        if path != ["Workbook"] {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            assert!(
                old.open_stream(&refs).unwrap() == new.open_stream(&refs).unwrap(),
                "changed {path:?}"
            );
        }
    }
    check_catalog(&after);
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened
            .workbook()
            .worksheet("Sheet1")
            .unwrap()
            .unwrap()
            .cell_value(litchi_xls::cell_values::Reference::new(3, 3).unwrap())
            .unwrap(),
        Some(CellValue::String("Reviewed ordinary worksheet".into()))
    );
    let before_wb = workbook(&bytes);
    let after_wb = workbook(&after);
    let (before_start, before) = macro_stream(&before_wb);
    let (after_start, after_macro) = macro_stream(&after_wb);
    let delta = after_start as i64 - before_start as i64;
    let before = Records::new(before, Limits::default())
        .unwrap()
        .map(|r| r.unwrap())
        .collect::<Vec<_>>();
    let after = Records::new(after_macro, Limits::default())
        .unwrap()
        .map(|r| r.unwrap())
        .collect::<Vec<_>>();
    assert_eq!(before.len(), after.len());
    for (a, b) in before.iter().zip(&after) {
        assert_eq!(a.kind(), b.kind());
        let mut expected = a.payload().to_vec();
        if a.kind().get() == 0x020b {
            for (i, value) in expected[12..].chunks_exact_mut(4).enumerate() {
                let offset = u32::from_le_bytes(value.try_into().unwrap()) as usize;
                assert_eq!(
                    u16::from_le_bytes(before_wb[offset..offset + 2].try_into().unwrap()),
                    if i == 0 { 0x55 } else { 0xd7 }
                );
                let moved = (offset as i64 + delta) as usize;
                assert_eq!(
                    u16::from_le_bytes(after_wb[moved..moved + 2].try_into().unwrap()),
                    if i == 0 { 0x55 } else { 0xd7 }
                );
                value.copy_from_slice(&(moved as u32).to_le_bytes());
            }
        }
        assert_eq!(b.payload(), expected, "macro record {:04x}", a.kind().get());
    }
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
