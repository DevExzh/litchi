//! Native BIFF calculation and workbook-window metadata; no recalculation.
use litchi_core::sheet::{Cell as _, WorkbookTrait, Worksheet as _};
use litchi_xls::{Package, raw::Reader};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_calculation_windows_and_selections_retain_exact_saved_state() {
    use litchi_xls::{CalculationMode, ReferenceMode, view::PaneType};
    for (name, engine, ids, active, active_name, geometry) in [
        (
            "excel-values-formulas-chart.xls",
            967461,
            vec![2, 1, 3],
            2,
            "Sheet2",
            (2380, 600, 28040, 16240),
        ),
        (
            "excel-interaction-print.xls",
            967461,
            vec![1],
            0,
            "Sheet1",
            (2380, 600, 28040, 16240),
        ),
        (
            "excel-metadata.xls",
            977461,
            vec![1, 2, 3, 4, 5],
            4,
            "Print_015",
            (3120, 1600, 27240, 15140),
        ),
        (
            "excel-date1904-breaks.xls",
            977461,
            vec![1],
            0,
            "Sheet1",
            (2320, 600, 28040, 16140),
        ),
    ] {
        let bytes = native(name);
        let b = Reader::new(Cursor::new(&bytes)).unwrap();
        let calc = b.calculation();
        assert!(calc.full_precision());
        assert!(!calc.force_full_calculation());
        assert_eq!(calc.recalculation_engine_id(), Some(engine));
        let mtr = calc.multithreaded_calculation().unwrap();
        assert!(mtr.enabled());
        assert_eq!(mtr.user_thread_count(), None);
        for name in b.worksheet_names() {
            let s = b.worksheet_by_name(name).unwrap();
            let c = s.calculation();
            assert_eq!(c.mode(), CalculationMode::Automatic);
            assert_eq!(c.maximum_iterations(), 100);
            assert_eq!(c.iteration_delta(), 0.001);
            assert_eq!(c.reference_mode(), ReferenceMode::A1);
            assert!(!c.iteration_enabled() && !c.formulas_pending_recalculation());
            assert!(c.recalculate_before_save());
        }
        let views = b.workbook_view();
        assert_eq!(views.sheet_ids(), ids);
        assert_eq!(views.windows().len(), 1);
        let w = views.primary_window().unwrap();
        assert_eq!(w.active_sheet_index(), active);
        assert_eq!(b.active_worksheet().unwrap().name(), active_name);
        assert_eq!(
            (
                w.horizontal_position_twips(),
                w.vertical_position_twips(),
                w.width_twips(),
                w.height_twips()
            ),
            geometry
        );
        assert_eq!(
            (
                w.first_visible_sheet_index(),
                w.selected_sheet_count(),
                w.sheet_tab_ratio_per_mille()
            ),
            (0, 1, 600)
        );
        assert!(
            w.shows_horizontal_scrollbar()
                && w.shows_vertical_scrollbar()
                && w.shows_sheet_tabs()
                && w.groups_dates_in_autofilter()
        );
        assert!(!w.hidden() && !w.minimized() && !w.very_hidden());
        if name == "excel-values-formulas-chart.xls" {
            assert_eq!(b.active_sheet_index(), 1); // ordinary-worksheet ordinal; chartsheet is tab 0
            let s = b.worksheet_by_name("Sheet1").unwrap();
            let v = s.worksheet_view().unwrap();
            let p = v.pane().unwrap();
            assert_eq!(
                (
                    p.horizontal_split(),
                    p.vertical_split(),
                    p.bottom_pane_top_row(),
                    p.right_pane_left_column()
                ),
                (1, 1, 1, 1)
            );
            assert_eq!(p.active_pane(), PaneType::LowerRight);
            assert_eq!(
                v.selections()
                    .iter()
                    .map(|s| (s.pane(), s.active_row(), s.active_column()))
                    .collect::<Vec<_>>(),
                [
                    (PaneType::UpperLeft, 0, 0),
                    (PaneType::UpperRight, 0, 1),
                    (PaneType::LowerLeft, 1, 0),
                    (PaneType::LowerRight, 1, 1)
                ]
            );
        } else if name == "excel-date1904-breaks.xls" {
            let v = b.xls_worksheet(0).unwrap().worksheet_view().unwrap();
            assert_eq!(v.selections().len(), 1);
            assert_eq!(
                (
                    v.selections()[0].active_row(),
                    v.selections()[0].active_column()
                ),
                (5, 1)
            );
        }
        assert!(Package::from_bytes(&bytes).unwrap().to_bytes().unwrap() == bytes);
    }
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut f = litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .unwrap();
    f.list_streams()
        .into_iter()
        .map(|p| {
            let refs = p.iter().map(String::as_str).collect::<Vec<_>>();
            let b = f.open_stream(&refs).unwrap();
            (p, b)
        })
        .collect()
}
fn metadata_records(stream: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut pos = 0;
    let mut result = Vec::new();
    while pos + 4 <= stream.len() {
        let kind = u16::from_le_bytes(stream[pos..pos + 2].try_into().unwrap());
        let n = usize::from(u16::from_le_bytes(
            stream[pos + 2..pos + 4].try_into().unwrap(),
        ));
        if matches!(
            kind,
            0x000c
                ..=0x0011
                    | 0x003d
                    | 0x005e
                    | 0x005f
                    | 0x013d
                    | 0x01c1
                    | 0x089a
                    | 0x08a3
                    | 0x023e
                    | 0x0041
                    | 0x001d
                    | 0x00a0
        ) {
            result.push((kind, stream[pos + 4..pos + 4 + n].to_vec()));
        }
        pos += 4 + n;
    }
    result
}
#[test]
fn native_cell_edit_preserves_calculation_view_records_and_other_cells() {
    let bytes = native("excel-values-formulas-chart.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let before = Reader::new(Cursor::new(&bytes)).unwrap();
    let mut edit = p.edit().unwrap();
    edit.set_numeric(
        "Sheet1",
        litchi_xls::cell_values::Reference::new(1, 1).unwrap(),
        123.,
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(f.path()).unwrap();
    let after = std::fs::read(f.path()).unwrap();
    let read = Reader::new(Cursor::new(&after)).unwrap();
    assert_eq!(
        read.xls_worksheet(0)
            .unwrap()
            .get_cell(1, 1)
            .unwrap()
            .value(),
        &litchi_core::sheet::CellValue::Float(123.)
    );
    assert_eq!(before.calculation(), read.calculation());
    assert_eq!(before.workbook_view(), read.workbook_view());
    for name in before.worksheet_names() {
        let a = before.worksheet_by_name(name).unwrap();
        let b = read.worksheet_by_name(name).unwrap();
        assert_eq!(a.calculation(), b.calculation());
        assert_eq!(a.worksheet_views(), b.worksheet_views());
        for c in a.cells() {
            let c = c.unwrap();
            if name == "Sheet1" && (c.row(), c.column()) == (2, 2) {
                continue;
            }
            let other = b.get_cell(c.row() - 1, c.column() - 1).unwrap();
            assert_eq!(c.value(), other.value());
            assert_eq!(c.formula(), other.formula());
        }
    }
    let a = streams(&bytes);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if name == ["Workbook"] {
            assert_eq!(metadata_records(&data), metadata_records(&b[&name]));
        } else {
            assert!(b[&name] == data);
        }
    }
    assert!(
        Package::from_bytes(&after)
            .unwrap()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
