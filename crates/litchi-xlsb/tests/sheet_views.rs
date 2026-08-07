//! Tests for XLSB worksheet sheet views (BrtBeginWsView / BrtPane / BrtSel).
//!
//! Covers the new-workbook writer (default, freeze panes, explicit views),
//! the read path against an Excel-produced fixture, and byte preservation of
//! untouched worksheet streams.

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI};
use litchi_sheet::view::{
    Color, Display, Mode, Pane, Position, Scale, Selection, Split, State, View, Window, Zoom,
};
use litchi_sheet::{Cell, Rect};
use litchi_xlsb::Workbook;
use litchi_xlsb::raw::{Kind, Records, kind};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

/// Resolve a repository-relative fixture path.
fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn workbook_bytes(configure: impl FnOnce(&mut MutableWorksheet)) -> Vec<u8> {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Sheet1");
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

fn records(data: &[u8]) -> Vec<(Kind, Vec<u8>)> {
    let mut result = Vec::new();
    for record in Records::new(data) {
        let record = record.unwrap();
        result.push((record.kind(), record.payload().to_vec()));
    }
    result
}

fn sheet_records(package_bytes: &[u8]) -> Vec<(Kind, Vec<u8>)> {
    records(&part_blob(package_bytes, "/xl/worksheets/sheet1.bin"))
}

fn cell(reference: &str) -> Cell {
    Cell::from_a1(reference).unwrap()
}

fn rect(reference: &str) -> Rect {
    Rect::from_a1(reference).unwrap()
}

fn scale(value: u16) -> Scale {
    Scale::new(value).unwrap()
}

fn split(value: f64) -> Split {
    Split::new(value).unwrap()
}

fn first_view(package_bytes: &[u8]) -> View {
    let workbook = Workbook::new(Cursor::new(package_bytes)).unwrap();
    let worksheet = workbook.worksheet(0).unwrap();
    let views = worksheet.views();
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
        .find(|(record_kind, _)| *record_kind == kind::BEGIN_WS_VIEW)
        .expect("worksheet contains BrtBeginWsView");
    assert_eq!(view.1, legacy);
    assert!(
        !sheet
            .iter()
            .any(|(record_kind, _)| *record_kind == kind::PANE)
    );
    assert!(
        !sheet
            .iter()
            .any(|(record_kind, _)| *record_kind == kind::SEL)
    );

    let view = first_view(&bytes);
    assert!(view.tab_selected);
    assert_eq!(view.zoom.current, scale(100));
    assert_eq!(view.mode, Mode::Normal);
    assert_eq!(view.origin, cell("A1"));
    assert_eq!(view.color, Color::DEFAULT);
    assert!(view.pane.is_none());
    assert!(view.selections.is_empty());
}

#[test]
fn freeze_panes_round_trip() {
    let bytes = workbook_bytes(|sheet| sheet.freeze_panes(2, 1).unwrap());

    let sheet = sheet_records(&bytes);
    let pane = sheet
        .iter()
        .find(|(record_kind, _)| *record_kind == kind::PANE)
        .expect("worksheet contains BrtPane");
    assert_eq!(pane.1.len(), 29);
    let selections: Vec<_> = sheet
        .iter()
        .filter(|(record_kind, _)| *record_kind == kind::SEL)
        .collect();
    assert_eq!(selections.len(), 1);

    let view = first_view(&bytes);
    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, State::Frozen);
    assert_eq!(pane.horizontal, Some(split(1.0)));
    assert_eq!(pane.vertical, Some(split(2.0)));
    assert_eq!(pane.top_left, cell("B3"));
    assert_eq!(pane.position, Position::BottomRight);
    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.position(), Position::BottomRight);
    assert_eq!(selection.active_cell(), cell("B3"));
    assert_eq!(selection.active_range(), 0);
    assert_eq!(selection.ranges(), [Rect::single(cell("B3"))]);
}

#[test]
fn freeze_rows_only_uses_bottom_left_pane() {
    let bytes = workbook_bytes(|sheet| sheet.freeze_panes(3, 0).unwrap());
    let view = first_view(&bytes);
    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, State::Frozen);
    assert_eq!(pane.horizontal, None);
    assert_eq!(pane.vertical, Some(split(3.0)));
    assert_eq!(pane.top_left, cell("A4"));
    assert_eq!(pane.position, Position::BottomLeft);
}

#[test]
fn unfreeze_panes_removes_pane_records_and_preserves_unrelated_view_state() {
    let bytes = workbook_bytes(|sheet| {
        let mut view = View::default();
        view.window = Window::new(7);
        view.mode = Mode::PageLayout;
        view.display.grid_lines = false;
        view.display.right_to_left = true;
        view.zoom = Zoom {
            current: scale(150),
            normal: Some(scale(75)),
            page_layout: Some(scale(60)),
            page_break_preview: Some(scale(200)),
        };
        view.origin = cell("C7");
        sheet.set_view(view);
        sheet.freeze_panes(1, 1).unwrap();
        sheet.unfreeze_panes();
    });
    let sheet = sheet_records(&bytes);
    assert!(
        !sheet
            .iter()
            .any(|(record_kind, _)| *record_kind == kind::PANE)
    );
    let view = first_view(&bytes);
    assert_eq!(view.window, Window::new(7));
    assert_eq!(view.mode, Mode::PageLayout);
    assert!(!view.display.grid_lines);
    assert!(view.display.right_to_left);
    assert_eq!(view.zoom.current, scale(150));
    assert_eq!(view.zoom.normal, Some(scale(75)));
    assert_eq!(view.zoom.page_layout, Some(scale(60)));
    assert_eq!(view.zoom.page_break_preview, Some(scale(200)));
    assert_eq!(view.origin, cell("C7"));
    assert!(view.pane.is_none());
}

#[test]
fn explicit_sheet_view_round_trip() {
    let bytes = workbook_bytes(|sheet| {
        let mut view = View::default();
        view.tab_selected = false;
        view.display.grid_lines = false;
        view.display.right_to_left = true;
        view.mode = Mode::PageBreakPreview;
        view.origin = cell("C7");
        view.zoom = Zoom {
            current: scale(150),
            normal: Some(scale(75)),
            ..view.zoom
        };
        view.pane = Some(Pane {
            horizontal: Some(split(2.0)),
            vertical: None,
            top_left: cell("C1"),
            position: Position::TopRight,
            state: State::FrozenSplit,
        });
        view.selections = vec![
            Selection::new(
                Position::TopRight,
                cell("D5"),
                1,
                vec![rect("A1:B2"), Rect::single(cell("D5"))],
            )
            .unwrap(),
        ];
        sheet.set_view(view);
    });

    let view = first_view(&bytes);
    assert!(!view.tab_selected);
    assert!(!view.display.grid_lines);
    assert!(view.display.right_to_left);
    assert_eq!(view.mode, Mode::PageBreakPreview);
    assert_eq!(view.origin, cell("C7"));
    assert_eq!(view.zoom.current, scale(150));
    assert_eq!(view.zoom.normal, Some(scale(75)));

    let pane = view.pane.expect("view has a pane");
    assert_eq!(pane.state, State::FrozenSplit);
    assert_eq!(pane.horizontal, Some(split(2.0)));
    assert_eq!(pane.vertical, None);
    assert_eq!(pane.top_left, cell("C1"));
    assert_eq!(pane.position, Position::TopRight);

    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.position(), Position::TopRight);
    assert_eq!(selection.active_cell(), cell("D5"));
    assert_eq!(selection.active_range(), 1);
    assert_eq!(
        selection.ranges(),
        [rect("A1:B2"), Rect::single(cell("D5"))]
    );
}

#[test]
fn freeze_panes_updates_canonical_view_without_losing_non_pane_state() {
    let mut sheet = MutableWorksheet::new("Sheet1");
    let mut view = View::default();
    view.window = Window::new(9);
    view.color = Color::new(64).unwrap();
    view.mode = Mode::PageLayout;
    view.display = Display {
        window_protection: true,
        show_formulas: true,
        grid_lines: false,
        row_column_headers: false,
        zero_values: false,
        right_to_left: true,
        ruler: false,
        outline_symbols: false,
        default_grid_color: false,
        white_space: false,
    };
    view.zoom = Zoom {
        current: scale(150),
        normal: Some(scale(75)),
        page_layout: Some(scale(60)),
        page_break_preview: Some(scale(200)),
    };
    view.origin = cell("C7");
    sheet.set_view(view);
    sheet.freeze_panes(1, 1).unwrap();
    let view = sheet.view().expect("freeze panes creates a canonical view");
    assert_eq!(view.window, Window::new(9));
    assert_eq!(view.color, Color::new(64).unwrap());
    assert_eq!(view.mode, Mode::PageLayout);
    assert_eq!(
        view.display,
        Display {
            window_protection: true,
            show_formulas: true,
            grid_lines: false,
            row_column_headers: false,
            zero_values: false,
            right_to_left: true,
            ruler: false,
            outline_symbols: false,
            default_grid_color: false,
            white_space: false,
        }
    );
    assert_eq!(view.zoom.current, scale(150));
    assert_eq!(view.zoom.normal, Some(scale(75)));
    assert_eq!(view.zoom.page_layout, Some(scale(60)));
    assert_eq!(view.zoom.page_break_preview, Some(scale(200)));
    assert_eq!(view.origin, cell("C7"));

    let pane = view.pane.as_ref().expect("freeze panes installs a pane");
    assert_eq!(pane.position, Position::BottomRight);
    assert_eq!(pane.state, State::Frozen);
    assert_eq!(pane.horizontal, Some(split(1.0)));
    assert_eq!(pane.vertical, Some(split(1.0)));
    assert_eq!(pane.top_left, cell("B2"));
}

#[test]
fn reads_excel_fixture_views_and_selections() {
    let path = fixture("ooxml/xlsb/Simple.xlsb");
    let workbook = Workbook::new(File::open(&path).unwrap()).unwrap();

    let first = workbook.worksheet(0).unwrap();
    let views = first.views();
    assert_eq!(views.len(), 1);
    let view = &views[0];
    assert!(view.tab_selected);
    assert!(view.display.grid_lines);
    assert_eq!(view.zoom.current, scale(100));
    assert_eq!(view.mode, Mode::Normal);
    assert!(view.pane.is_none());
    assert_eq!(view.selections.len(), 1);
    let selection = &view.selections[0];
    assert_eq!(selection.position(), Position::TopLeft);
    assert_eq!(selection.active_cell(), cell("A1"));
    assert_eq!(selection.active_range(), 0);
    assert_eq!(selection.ranges(), [Rect::single(cell("A1"))]);

    let second = workbook.worksheet(1).unwrap();
    let views = second.views();
    assert_eq!(views.len(), 1);
    assert!(!views[0].tab_selected);
}

#[test]
fn untouched_worksheet_stream_round_trips_byte_identical() {
    let path = fixture("ooxml/xlsb/Simple.xlsb");
    let original = std::fs::read(&path).unwrap();
    let workbook = Workbook::new(Cursor::new(&original)).unwrap();
    // Force the read path across every worksheet, then save unmodified.
    for index in 0..workbook.worksheet_names().len() {
        let worksheet = workbook.worksheet(index).unwrap();
        assert_eq!(worksheet.views().len(), 1);
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
