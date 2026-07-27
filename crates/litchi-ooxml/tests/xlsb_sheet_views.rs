//! Tests for XLSB worksheet sheet views (BrtBeginWsView / BrtPane / BrtSel).
//!
//! Covers the new-workbook writer (default, freeze panes, explicit views),
//! the read path against an Excel-produced fixture, and byte preservation of
//! untouched worksheet streams.

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_ooxml::xlsb::writer::{MutableXlsbWorksheet, XlsbWorkbookWriter};
use litchi_ooxml::xlsb::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
    XlsbRecord, XlsbWorkbook,
};
use litchi_opc::{OpcPackage, PackURI};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

const BEGIN_WS_VIEW: u16 = 0x0089;
const PANE: u16 = 0x0097;
const SEL: u16 = 0x0098;

/// Resolve a repository-relative fixture path.
fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn workbook_bytes(configure: impl FnOnce(&mut MutableXlsbWorksheet)) -> Vec<u8> {
    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "views");
    configure(&mut sheet);
    workbook.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    output.into_inner()
}

fn part_blob(package_bytes: &[u8], path: &str) -> Vec<u8> {
    let package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    package
        .get_part(&PackURI::new(path).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn records(data: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut cursor = Cursor::new(data);
    let mut result = Vec::new();
    while cursor.position() < data.len() as u64 {
        let record = XlsbRecord::read(&mut cursor).unwrap();
        result.push((record.header.record_type, record.data.to_vec()));
    }
    result
}

fn sheet_records(package_bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    records(&part_blob(package_bytes, "/xl/worksheets/sheet1.bin"))
}

fn first_view(package_bytes: &[u8]) -> SheetView {
    let workbook = XlsbWorkbook::new(Cursor::new(package_bytes)).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    let views = worksheet.sheet_views();
    assert_eq!(views.len(), 1);
    views[0].clone()
}

#[test]
fn default_view_matches_legacy_writer_bytes() {
    // The 30-byte default BrtBeginWsView payload the crate has always emitted.
    let legacy = [
        0xDC, 0x03, // flags
        0x00, 0x00, 0x00, 0x00, // xlView
        0x00, 0x00, 0x00, 0x00, // rwTop
        0x00, 0x00, 0x00, 0x00, // colLeft
        0x40, // icvHdr
        0x00, // reserved2
        0x00, 0x00, // reserved3
        0x64, 0x00, // wScale
        0x00, 0x00, // wScaleNormal
        0x00, 0x00, // wScaleSLV
        0x00, 0x00, // wScalePLV
        0x00, 0x00, 0x00, 0x00, // iWbkView
    ];
    let bytes = workbook_bytes(|_| {});
    let sheet = sheet_records(&bytes);
    let view = sheet
        .iter()
        .find(|(kind, _)| *kind == BEGIN_WS_VIEW)
        .expect("worksheet contains BrtBeginWsView");
    assert_eq!(view.1, legacy);
    assert!(!sheet.iter().any(|(kind, _)| *kind == PANE));
    assert!(!sheet.iter().any(|(kind, _)| *kind == SEL));

    let view = first_view(&bytes);
    assert_eq!(view.tab_selected, Some(true));
    assert_eq!(view.zoom_scale, Some(100));
    assert_eq!(view.view_type, Some(SheetViewType::Normal));
    assert_eq!(view.top_left_cell.as_deref(), Some("A1"));
    assert!(view.pane.is_none());
    assert!(view.selections.is_empty());
}

#[test]
fn freeze_panes_round_trip() {
    let bytes = workbook_bytes(|sheet| sheet.freeze_panes(2, 1));

    let sheet = sheet_records(&bytes);
    let pane = sheet
        .iter()
        .find(|(kind, _)| *kind == PANE)
        .expect("worksheet contains BrtPane");
    assert_eq!(pane.1.len(), 29);
    let selections: Vec<_> = sheet.iter().filter(|(kind, _)| *kind == SEL).collect();
    assert_eq!(selections.len(), 1);

    let view = first_view(&bytes);
    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, Some(SheetPaneState::Frozen));
    assert_eq!(pane.x_split, Some(1.0));
    assert_eq!(pane.y_split, Some(2.0));
    assert_eq!(pane.top_left_cell.as_deref(), Some("B3"));
    assert_eq!(pane.active_pane, Some(SheetPanePosition::BottomRight));
    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.pane, Some(SheetPanePosition::BottomRight));
    assert_eq!(selection.active_cell.as_deref(), Some("B3"));
    assert_eq!(selection.sqref.as_deref(), Some("B3"));
}

#[test]
fn freeze_rows_only_uses_bottom_left_pane() {
    let bytes = workbook_bytes(|sheet| sheet.freeze_panes(3, 0));
    let view = first_view(&bytes);
    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, Some(SheetPaneState::Frozen));
    assert_eq!(pane.x_split, None);
    assert_eq!(pane.y_split, Some(3.0));
    assert_eq!(pane.top_left_cell.as_deref(), Some("A4"));
    assert_eq!(pane.active_pane, Some(SheetPanePosition::BottomLeft));
}

#[test]
fn unfreeze_panes_removes_pane_records() {
    let bytes = workbook_bytes(|sheet| {
        sheet.freeze_panes(1, 1);
        sheet.unfreeze_panes();
    });
    let sheet = sheet_records(&bytes);
    assert!(!sheet.iter().any(|(kind, _)| *kind == PANE));
}

#[test]
fn explicit_sheet_view_round_trip() {
    let bytes = workbook_bytes(|sheet| {
        sheet.set_sheet_view(SheetView {
            tab_selected: Some(false),
            show_grid_lines: Some(false),
            right_to_left: Some(true),
            view_type: Some(SheetViewType::PageBreakPreview),
            top_left_cell: Some("C7".to_string()),
            zoom_scale: Some(150),
            zoom_scale_normal: Some(75),
            pane: Some(SheetPane {
                x_split: Some(2.0),
                y_split: None,
                top_left_cell: Some("C1".to_string()),
                active_pane: Some(SheetPanePosition::TopRight),
                state: Some(SheetPaneState::FrozenSplit),
            }),
            selections: vec![SheetSelection {
                pane: Some(SheetPanePosition::TopRight),
                active_cell: Some("D5".to_string()),
                active_cell_id: Some(1),
                sqref: Some("A1:B2 D5".to_string()),
            }],
            ..SheetView::default()
        });
    });

    let view = first_view(&bytes);
    assert_eq!(view.tab_selected, Some(false));
    assert_eq!(view.show_grid_lines, Some(false));
    assert_eq!(view.right_to_left, Some(true));
    assert_eq!(view.view_type, Some(SheetViewType::PageBreakPreview));
    assert_eq!(view.top_left_cell.as_deref(), Some("C7"));
    assert_eq!(view.zoom_scale, Some(150));
    assert_eq!(view.zoom_scale_normal, Some(75));

    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, Some(SheetPaneState::FrozenSplit));
    assert_eq!(pane.x_split, Some(2.0));
    assert_eq!(pane.y_split, None);
    assert_eq!(pane.top_left_cell.as_deref(), Some("C1"));
    assert_eq!(pane.active_pane, Some(SheetPanePosition::TopRight));

    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.pane, Some(SheetPanePosition::TopRight));
    assert_eq!(selection.active_cell.as_deref(), Some("D5"));
    assert_eq!(selection.active_cell_id, Some(1));
    assert_eq!(selection.sqref.as_deref(), Some("A1:B2 D5"));
}

#[test]
fn freeze_panes_conflict_with_explicit_pane_fails() {
    let mut workbook = XlsbWorkbookWriter::new();
    let mut sheet = MutableXlsbWorksheet::new("Sheet1");
    sheet.set_sheet_view(SheetView {
        pane: Some(SheetPane {
            state: Some(SheetPaneState::Split),
            ..SheetPane::default()
        }),
        ..SheetView::default()
    });
    sheet.freeze_panes(1, 1);
    workbook.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    assert!(workbook.save(&mut output).is_err());
}

#[test]
fn reads_excel_fixture_views_and_selections() {
    let path = fixture("test-data/ooxml/xlsb/Simple.xlsb");
    let workbook = XlsbWorkbook::new(File::open(&path).unwrap()).unwrap();

    let first = workbook.worksheet(0).unwrap();
    let views = first.sheet_views();
    assert_eq!(views.len(), 1);
    let view = &views[0];
    assert_eq!(view.tab_selected, Some(true));
    assert_eq!(view.show_grid_lines, Some(true));
    assert_eq!(view.zoom_scale, Some(100));
    assert_eq!(view.view_type, Some(SheetViewType::Normal));
    assert!(view.pane.is_none());
    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.pane, Some(SheetPanePosition::TopLeft));
    assert_eq!(selection.active_cell.as_deref(), Some("A1"));
    assert_eq!(selection.active_cell_id, Some(0));
    assert_eq!(selection.sqref.as_deref(), Some("A1"));

    let second = workbook.worksheet(1).unwrap();
    let views = second.sheet_views();
    assert_eq!(views.len(), 1);
    assert_eq!(views[0].tab_selected, Some(false));
}

#[test]
fn untouched_worksheet_stream_round_trips_byte_identical() {
    let path = fixture("test-data/ooxml/xlsb/Simple.xlsb");
    let original = std::fs::read(&path).unwrap();
    let workbook = XlsbWorkbook::new(Cursor::new(&original)).unwrap();
    // Force the read path across every worksheet, then save unmodified.
    for index in 0..workbook.worksheet_names().len() {
        let worksheet = workbook.worksheet(index).unwrap();
        assert_eq!(worksheet.sheet_views().len(), 1);
    }
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let saved = output.into_inner();

    for name in [
        "/xl/worksheets/sheet1.bin",
        "/xl/worksheets/sheet2.bin",
        "/xl/worksheets/sheet3.bin",
    ] {
        assert_eq!(
            part_blob(&original, name),
            part_blob(&saved, name),
            "{name} must round-trip byte-identically"
        );
    }
}
