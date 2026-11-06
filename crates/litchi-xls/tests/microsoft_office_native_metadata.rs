//! Semantic reads and source-preserving operations on native Excel BIFF8 files.

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::{Package, PivotItemType, raw::Reader};
use std::collections::BTreeMap;
use std::io::Cursor;

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}

fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).unwrap();
            (path, data)
        })
        .collect()
}

fn records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut result = Vec::new();
    let mut pos = 0;
    while pos < bytes.len() {
        let kind = u16::from_le_bytes(bytes[pos..pos + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            bytes[pos + 2..pos + 4].try_into().unwrap(),
        ));
        result.push((kind, bytes[pos + 4..pos + 4 + len].to_vec()));
        pos += 4 + len;
    }
    result
}

fn assert_other_streams(before: &[u8], after: &[u8]) {
    let before = streams(before);
    let after = streams(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in &before {
        if path != &["Workbook"] {
            assert!(after[path] == *data, "changed {path:?}");
        }
    }
}

fn number(sheet: &litchi_xls::Worksheet, row: u32, column: u32, expected: f64) {
    assert_eq!(
        sheet.get_cell(row, column).unwrap().value(),
        &CellValue::Float(expected)
    );
}

#[test]
fn native_pivot_items_distinguish_data_from_automatic_subtotals() {
    for file in ["excel-visual-analytics.xls", "excel-pivot-controls.xls"] {
        let bytes = native(file);
        let reader = Reader::new(Cursor::new(&bytes)).unwrap();
        let sheet = reader.xls_worksheet(0).unwrap();
        let table = sheet.pivot_table("PivotTable1").unwrap();
        for field in [1, 2] {
            let items = &table.fields[field].items;
            assert_eq!(items.len(), 3);
            assert_eq!(items[0].item_type, PivotItemType::Data);
            assert_eq!(items[1].item_type, PivotItemType::Data);
            assert_eq!(items[2].item_type, PivotItemType::Default);
            let expected = if field == 1 {
                [0, 1, 0xffff]
            } else {
                [1, 0, 0xffff]
            };
            assert_eq!(
                [
                    items[0].cache_index,
                    items[1].cache_index,
                    items[2].cache_index
                ],
                expected
            );
        }
        assert_eq!(table.row_lines.last().unwrap().item_type, 13);
        assert_eq!(table.column_lines.last().unwrap().item_type, 13);
    }
}

#[test]
fn native_pivot_cache_rows_axes_and_aggregations_retain_saved_totals() {
    use litchi_xls::{
        PivotAxis, PivotAxisField, PivotCacheDateTime, PivotCacheItem as Item, PivotFunction,
        PivotPageSelection,
    };
    for (file, last_row, last_col, total) in [
        ("excel-visual-analytics.xls", 6, 3, 175.5),
        ("excel-pivot-controls.xls", 5, 2, 9.5),
    ] {
        let bytes = native(file);
        let book = Reader::new(Cursor::new(&bytes)).unwrap();
        assert_eq!(book.pivot_caches().len(), 1);
        assert_eq!(book.pivot_cache_stream_ids(), [1]);
        let sheet = book.xls_worksheet(0).unwrap();
        let table = sheet.pivot_table("PivotTable1").unwrap();
        let cache = book.pivot_cache_for_table(table).unwrap();
        assert_eq!(
            cache.fields().iter().map(|f| f.name()).collect::<Vec<_>>(),
            ["Date", "Region", "Product", "Revenue", "Units", "Status"]
        );
        assert_eq!(cache.rows().len(), 8);
        for (row, (month, day, region, product, revenue, units, status)) in
            cache.rows().iter().zip([
                (1, 15, "North", "Widget", 12.5, 3.0, "Open"),
                (1, 20, "South", "Widget", 25.0, 5.0, "Closed"),
                (2, 10, "North", "Gadget", 9.5, 2.0, "Open"),
                (2, 18, "South", "Gadget", 48.0, 4.0, "Hold"),
                (3, 3, "North", "Widget", 16.5, 3.0, "Open"),
                (3, 21, "South", "Gadget", 20.0, 4.0, "Closed"),
                (4, 2, "North", "Gadget", 14.0, 2.0, "Hold"),
                (4, 18, "South", "Widget", 30.0, 6.0, "Open"),
            ])
        {
            assert_eq!(
                row,
                &[
                    Item::DateTime(PivotCacheDateTime::try_new(2026, month, day, 0, 0, 0).unwrap()),
                    Item::String(region.into()),
                    Item::String(product.into()),
                    Item::Number(revenue),
                    Item::Number(units),
                    Item::String(status.into()),
                ]
            );
        }
        assert_eq!(
            (
                table.view.first_row,
                table.view.first_col,
                table.view.last_row,
                table.view.last_col
            ),
            (2, 0, last_row, last_col)
        );
        assert_eq!(table.row_fields, [PivotAxisField::Field(1)]);
        assert_eq!(table.column_fields, [PivotAxisField::Field(2)]);
        assert_eq!(table.fields[1].axis, PivotAxis::Row);
        assert_eq!(table.fields[2].axis, PivotAxis::Column);
        assert_eq!(table.fields[5].axis, PivotAxis::Page);
        assert_eq!(table.data_items.len(), 1);
        assert_eq!(table.data_items[0].source_field_index, 3);
        assert_eq!(table.data_items[0].function, PivotFunction::Sum);
        assert_eq!(table.data_items[0].name, "Sum of Revenue");
        assert_eq!(table.page_entries.len(), 1);
        assert_eq!(table.page_entries[0].field_index, 5);
        assert_eq!(table.page_entries[0].selection(), PivotPageSelection::All);
        number(sheet, u32::from(last_row), u32::from(last_col), total);
    }
}

#[test]
fn native_formula_owners_keep_cse_shared_dynamic_and_data_table_storage() {
    use litchi_xls::{DataTableInputCell as Input, DataTableKind};
    let bytes = native("excel-formula-owners.xls");
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = reader.xls_worksheet(0).unwrap();
    assert_eq!(sheet.array_formulas().len(), 1);
    let owner = sheet.array_formulas().next().unwrap();
    assert_eq!(
        (
            owner.range().first().row(),
            owner.range().first().col(),
            owner.range().last().row(),
            owner.range().last().col()
        ),
        (3, 2, 6, 2)
    );
    assert_eq!(
        owner.tokens(),
        [
            69, 3, 0, 6, 0, 0, 192, 0, 192, 69, 3, 0, 6, 0, 1, 192, 1, 192, 5
        ]
    );
    for (row, value) in (3..7).zip([20.0, 60.0, 120.0, 200.0]) {
        number(sheet, row, 2, value);
        let cell = sheet.get_cell(row, 2).unwrap();
        assert_eq!(cell.formula(), Some("=(A4:A7*B4:B7)"));
        assert!(cell.is_array_formula());
        assert_eq!(cell.array_formula(), Some(owner));
        assert_eq!(cell.formula_bytes(), Some([1, 3, 0, 2, 0].as_slice()));
    }
    for row in 3..103 {
        number(sheet, row, 3, f64::from((row + 1) * 2));
        let cell = sheet.get_cell(row, 3).unwrap();
        assert_eq!(cell.formula(), Some("=(ROW()*2)"));
        assert_eq!(cell.formula_metadata().unwrap().shared_formula(), row != 3);
        if row != 3 {
            let anchor = if row < 68 { 4 } else { 68 };
            assert_eq!(cell.formula_bytes(), Some([1, anchor, 0, 3, 0].as_slice()));
        }
    }
    let tables = sheet.data_tables();
    assert_eq!(tables.len(), 2);
    for (table, first_col, last_col) in [(tables[0], 9, 9), (tables[1], 12, 15)] {
        let range = table.range();
        assert_eq!(
            (
                range.first_row(),
                range.last_row(),
                range.first_col(),
                range.last_col()
            ),
            (3, 6, first_col, last_col)
        );
    }
    assert!(matches!(
        tables[0].kind(),
        DataTableKind::OneVariable {
            input: Input::Present { row: 3, col: 7 },
            ..
        }
    ));
    assert!(matches!(
        tables[1].kind(),
        DataTableKind::TwoVariable {
            row_input: Input::Present { row: 4, col: 7 },
            column_input: Input::Present { row: 3, col: 7 }
        }
    ));
    for (row, value) in (3..7).zip([7.0, 9.0, 11.0, 13.0]) {
        number(sheet, row, 9, value);
    }
    for (row, values) in (3..7).zip([
        [15.0, 25.0, 35.0, 45.0],
        [25.0, 45.0, 65.0, 85.0],
        [35.0, 65.0, 95.0, 125.0],
        [45.0, 85.0, 125.0, 165.0],
    ]) {
        for (col, value) in (12..16).zip(values) {
            number(sheet, row, col, value);
        }
    }
    let dynamic = reader.xls_worksheet(1).unwrap();
    assert_eq!(dynamic.array_formulas().len(), 2);
    for (row, col, value) in [
        (3, 0, 10.0),
        (3, 1, 12.0),
        (4, 0, 14.0),
        (4, 1, 16.0),
        (5, 0, 18.0),
        (5, 1, 20.0),
        (3, 3, 4.0),
        (3, 4, 30.0),
        (4, 3, 5.0),
        (4, 4, 40.0),
    ] {
        number(dynamic, row, col, value);
        let cell = dynamic.get_cell(row, col).unwrap();
        assert!(cell.is_array_formula());
        assert_eq!(cell.formula(), None); // Legacy renderer cannot express these UDF-backed owners.
        assert!(cell.formula_bytes().is_some());
        assert!(!cell.array_formula().unwrap().tokens().is_empty());
    }
    // Excel's XLS conversion stores LET as a #VALUE! token with its prior cache.
    let legacy_let = dynamic.get_cell(3, 5).unwrap();
    assert_eq!(legacy_let.formula(), Some("=#VALUE!"));
    assert_eq!(legacy_let.formula_bytes(), Some([0x1c, 0x0f].as_slice()));
    number(dynamic, 3, 5, 50.0);
}

#[test]
fn native_scenarios_keep_inputs_current_selection_and_flags_inert() {
    let bytes = native("excel-formula-owners.xls");
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = reader.xls_worksheet(0).unwrap();
    let manager = sheet.scenario_manager().unwrap();
    assert_eq!(manager.current_scenario(), Some(0));
    assert_eq!(manager.shown_scenario(), None);
    assert_eq!(manager.scenarios().len(), 2);
    for (scenario, (name, comment, values)) in manager.scenarios().iter().zip([
        ("BASELINE_009", "Native scenario: X=10, Y=2.", ["10", "2"]),
        ("ALTERNATE_009", "Native scenario: X=25, Y=4.", ["25", "4"]),
    ]) {
        assert_eq!(scenario.name(), name);
        assert_eq!(scenario.comment(), Some(comment));
        assert_eq!(scenario.creator(), Some("Ryker Zhu"));
        assert!(scenario.is_locked());
        assert!(!scenario.is_hidden());
        assert_eq!(scenario.cells().len(), 2);
        for (cell, (row, value)) in scenario.cells().iter().zip([3, 4].into_iter().zip(values)) {
            assert_eq!((cell.row(), cell.column(), cell.value()), (row, 7, value));
            assert!(!cell.is_deleted());
        }
    }
    for (row, value) in [(3, 10.0), (4, 2.0), (5, 25.0)] {
        number(sheet, row, 7, value);
    }
}

#[test]
fn native_print_setup_notes_and_hyperlinks_retain_exact_metadata() {
    use litchi_xls::{PrintComments, PrintOrder, PrintOrientation};
    let bytes = native("excel-interaction-print.xls");
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = reader.xls_worksheet(0).unwrap();
    let page = sheet.page_setup().unwrap();
    assert!(page.print_headers());
    assert!(page.print_gridlines());
    assert!(page.is_horizontally_centered());
    assert_eq!(
        (
            page.left_margin_inches(),
            page.right_margin_inches(),
            page.top_margin_inches(),
            page.bottom_margin_inches()
        ),
        (Some(0.7), Some(0.7), Some(0.5), Some(0.5))
    );
    let font = "&\"Aptos Narrow,Regular\"&K000000";
    assert_eq!(
        page.header(),
        format!("&L{font}HEADER_LEFT_002&C{font}Native Excel print fixture&R{font}&A")
    );
    assert_eq!(
        page.footer(),
        format!("&L{font}FOOTER_002&C{font}Page &P of &N&R{font}&F")
    );
    let setup = page.print_setup();
    assert_eq!(setup.orientation(), Some(PrintOrientation::Landscape));
    assert_eq!(setup.scale_percent(), Some(80));
    assert_eq!(setup.starting_page_number(), Some(3));
    assert_eq!((setup.fit_width_pages(), setup.fit_height_pages()), (1, 1));
    assert_eq!(setup.print_order(), PrintOrder::OverThenDown);
    assert_eq!(setup.comments(), PrintComments::AtEnd);
    assert_eq!(
        (
            setup.horizontal_resolution_dpi(),
            setup.vertical_resolution_dpi()
        ),
        (Some(0), Some(0))
    );
    assert_eq!(
        (setup.header_margin_inches(), setup.footer_margin_inches()),
        (0.3, 0.3)
    );
    assert_eq!(
        page.printer_driver_data(),
        [vec![
            2, 0, 0, 0, 1, 0, 0, 0, 80, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0
        ]]
    );
    assert_eq!(sheet.comments().len(), 2);
    for (note, (row, id, text)) in sheet.comments().iter().zip([
        (1, 3, "Ryker Zhu:\nLEGACY_NOTE_002: native Excel cell note."),
        (
            2,
            4,
            "Ryker Zhu:\nCOMMENT_COMMAND_002: comment inserted through Review ribbon.",
        ),
    ]) {
        assert_eq!((note.row(), note.column()), (row, 0));
        assert_eq!(note.author(), "Ryker Zhu");
        assert_eq!(note.text(), text);
        assert_eq!(note.identity().object_id(), id);
        assert_eq!(note.visibility(), litchi_xls::Visibility::Hidden);
    }
    let features = Package::from_bytes(&bytes)
        .unwrap()
        .sheet_features(0, litchi_xls::sheet_features::FeatureSet::HYPERLINKS)
        .unwrap();
    assert_eq!(features.hyperlinks.len(), 1);
    assert_eq!(
        features.hyperlinks[0].range(),
        sheet.hyperlinks()[0].range()
    );
    assert_eq!(
        sheet.hyperlinks()[0].tooltip(),
        Some("Native hyperlink screen tip 002")
    );
    assert_eq!(
        sheet.hyperlinks()[0].display_name(),
        Some("HYPERLINK_FIXTURE_002")
    );
    assert_eq!(
        sheet.hyperlinks()[0].address(),
        Some("https://example.com/litchi-excel-fixture")
    );
}

#[test]
fn native_scalar_crud_preserves_formula_records_macro_and_embedded_streams() {
    use litchi_xls::cell_values::{Reference, Value};
    for file in [
        "excel-formula-owners.xls",
        "excel-date1904-breaks.xls",
        "excel-visual-analytics.xls",
        "excel-recorded-macro.xls",
        "excel-recorded-macro.xlt",
        "excel-links-objects.xls",
    ] {
        let bytes = native(file);
        let source = Package::from_bytes(&bytes).unwrap();
        let mut edit = source.edit().unwrap();
        let at = Reference::new(120, 25).unwrap();
        edit.insert_cell("Sheet1", at, Value::Number(42.0)).unwrap();
        let added = edit.commit().unwrap();
        let after = added.package().to_bytes().unwrap();
        assert_other_streams(&bytes, &after);
        let original_streams = streams(&bytes);
        let changed_streams = streams(&after);
        let key = vec!["Workbook".to_string()];
        let formula_records = |stream: &[u8]| {
            records(stream)
                .into_iter()
                .filter(|(k, _)| [0x6, 0x221, 0x4bc, 0x236].contains(k))
                .collect::<Vec<_>>()
        };
        assert_eq!(
            formula_records(&original_streams[&key]),
            formula_records(&changed_streams[&key])
        );
        let reopened = Package::from_bytes(&after).unwrap();
        assert_eq!(
            reopened
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            Some(CellValue::Float(42.0))
        );
        let mut edit = reopened.edit().unwrap();
        edit.set_value("Sheet1", at, Value::Number(43.0)).unwrap();
        let updated = edit.commit().unwrap();
        assert_eq!(
            updated
                .package()
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            Some(CellValue::Float(43.0))
        );
        let mut edit = updated.package().edit().unwrap();
        edit.remove_cell("Sheet1", at).unwrap();
        let removed = edit.commit().unwrap();
        assert_eq!(
            removed
                .package()
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            None
        );
        let undone = reopened.apply(&added.patch().inverse()).unwrap();
        assert!(undone.package().to_bytes().unwrap() == bytes);
        assert!(source.to_bytes().unwrap() == bytes);
    }
}

#[test]
fn native_custom_views_distinguish_saved_hidden_and_visible_rows() {
    use litchi_xls::CustomViewHiddenRows;
    let bytes = native("excel-date1904-breaks.xls");
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = reader.xls_worksheet(0).unwrap();
    assert_eq!(reader.custom_views().len(), 2);
    assert_eq!(sheet.custom_views().len(), 2);
    for (name, hidden, guid) in [
        (
            "BASE_PRINT_012",
            false,
            [
                0x54, 0xfc, 0x5e, 0xee, 0xc0, 0xd2, 0x47, 0x65, 0xa2, 0xc4, 0xef, 0x07, 0x79, 0x24,
                0xc7, 0x5e,
            ],
        ),
        (
            "HIDDEN_DETAIL_012",
            true,
            [
                0xfb, 0xdb, 0x40, 0x13, 0xf5, 0xa9, 0x44, 0x81, 0xb0, 0x6f, 0x5f, 0x4f, 0x94, 0xc1,
                0x2b, 0x28,
            ],
        ),
    ] {
        let book_view = reader
            .custom_views()
            .iter()
            .find(|v| v.name() == name)
            .unwrap();
        assert_eq!(book_view.guid(), &guid);
        assert!(book_view.includes_print_settings());
        assert!(book_view.includes_hidden_rows_columns_and_filters());
        let view = sheet
            .custom_views()
            .iter()
            .find(|v| v.begin().guid() == &guid)
            .unwrap()
            .begin();
        assert_eq!(
            view.hidden_rows(),
            if hidden {
                CustomViewHiddenRows::Present
            } else {
                CustomViewHiddenRows::NotPresent
            }
        );
        assert_eq!(view.has_hidden_columns(), hidden);
        assert!(!view.suppresses_zeros());
        assert!(view.shows_page_breaks());
        assert!(view.has_print_area());
        assert_eq!((view.tab_id(), view.scale()), (1, 100));
    }
}

#[test]
fn native_1904_layout_and_page_break_edits_keep_unrelated_biff_records() {
    use litchi_xls::sheet_features::{FeatureSet, PageBreak};
    let bytes = native("excel-date1904-breaks.xls");
    let source = Package::from_bytes(&bytes).unwrap();
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    assert_eq!(reader.date_system(), litchi_xls::DateSystem::Excel1904);
    let sheet = reader.xls_worksheet(0).unwrap();
    for (row, value) in [
        (2, 0.0),
        (3, 1.0),
        (4, 43889.0),
        (5, 44848.5515625),
        (6, -1.0),
        (7, 0.2708333333333333),
    ] {
        // The reader keeps negative serials numeric while retaining their
        // date-format association; positive date/time serials are typed dates.
        let expected = if row == 6 {
            CellValue::Float(value)
        } else {
            CellValue::DateTime(value)
        };
        assert_eq!(sheet.get_cell(row, 1).unwrap().value(), &expected);
        assert!(sheet.is_date_time_formatted(row, 1));
    }
    let view = sheet.worksheet_view().unwrap();
    assert_eq!(view.selections().len(), 1);
    assert_eq!(
        (
            view.selections()[0].active_row(),
            view.selections()[0].active_column()
        ),
        (5, 1)
    );
    assert_eq!(
        sheet
            .row_layouts()
            .filter(|r| r.is_hidden())
            .map(|r| r.row())
            .collect::<Vec<_>>(),
        [13, 14]
    );
    let hidden = sheet.column_layout(9).unwrap();
    assert_eq!(
        (
            hidden.first_column(),
            hidden.last_column(),
            hidden.width_256ths(),
            hidden.is_hidden()
        ),
        (9, 9, 0, true)
    );
    assert_eq!(sheet.column_layout(0).unwrap().width_256ths(), 8405);
    let page = sheet.page_setup().unwrap();
    assert_eq!(
        (
            page.horizontal_page_breaks()[0].position(),
            page.vertical_page_breaks()[0].position()
        ),
        (19, 6)
    );
    let mut edit = source.edit_sheet_features().unwrap();
    let original = edit.features(0, FeatureSet::PAGE_BREAKS).unwrap();
    assert_eq!(
        original.row_breaks,
        [PageBreak {
            index: 19,
            min: 0,
            max: 255
        }]
    );
    assert_eq!(
        original.column_breaks,
        [PageBreak {
            index: 6,
            min: 0,
            max: 65535
        }]
    );
    let mut expected = original.row_breaks.clone();
    expected.push(PageBreak {
        index: 30,
        min: 0,
        max: 255,
    });
    edit.set_row_breaks(0, &expected).unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let result = reopened.sheet_features(0, FeatureSet::PAGE_BREAKS).unwrap();
    assert_eq!(result.row_breaks, expected);
    assert_eq!(result.column_breaks, original.column_breaks);
    assert_other_streams(&bytes, &after);
    let before = streams(&bytes);
    let changed = streams(&after);
    let key = vec!["Workbook".to_string()];
    let unrelated = |data: &[u8]| {
        records(data)
            .into_iter()
            .filter(|(kind, _)| ![0x001b, 0x0085, 0x020b].contains(kind))
            .collect::<Vec<_>>()
    };
    assert_eq!(unrelated(&before[&key]), unrelated(&changed[&key]));
    assert!(
        reopened
            .apply_sheet_features(&committed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );

    let mut edit = source.edit().unwrap();
    edit.set_row_hidden("Sheet1", 13, false).unwrap();
    edit.set_column_hidden("Sheet1", 9, false).unwrap();
    let layout = edit.commit().unwrap();
    let layout_bytes = layout.package().to_bytes().unwrap();
    let result = Reader::new(Cursor::new(&layout_bytes)).unwrap();
    assert!(
        !result
            .xls_worksheet(0)
            .unwrap()
            .row_layout(13)
            .unwrap()
            .is_hidden()
    );
    assert!(
        result
            .xls_worksheet(0)
            .unwrap()
            .row_layout(14)
            .unwrap()
            .is_hidden()
    );
    assert!(
        !result
            .xls_worksheet(0)
            .unwrap()
            .column_layout(9)
            .unwrap()
            .is_hidden()
    );
    assert_other_streams(&bytes, &layout_bytes);
    assert!(
        result
            .xls_worksheet(0)
            .unwrap()
            .row_block_index()
            .unwrap()
            .is_some()
    );
    let layout_streams = streams(&layout_bytes);
    // Row-layout publication rebuilds INDEX/DBCELL, including removing
    // unused native offset entries for empty rows. Validate that index
    // against its actual cell boundaries and compare all unrelated records.
    let untouched_layout = |data: &[u8]| {
        records(data)
            .into_iter()
            .filter(|(kind, payload)| {
                !([0x020b, 0x00d7].contains(kind)
                    || (*kind == 0x0208
                        && u16::from_le_bytes(payload[..2].try_into().unwrap()) == 13)
                    || (*kind == 0x007d
                        && u16::from_le_bytes(payload[..2].try_into().unwrap()) == 9))
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(
        untouched_layout(&before[&key]),
        untouched_layout(&layout_streams[&key])
    );
    assert!(
        layout
            .package()
            .apply(&layout.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}

#[test]
fn native_protected_notes_allow_exact_noop_and_refuse_changed_text_atomically() {
    use litchi_xls::{
        cell_values::{Reference, Selector},
        comments::Value,
    };
    let bytes = native("excel-interaction-print.xls");
    let source = Package::from_bytes(&bytes).unwrap();
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    assert!(reader.xls_worksheet(0).unwrap().protection().is_protected());
    let original = &reader.xls_worksheet(0).unwrap().comments()[0];
    let mut edit = source.edit_comments().unwrap();
    let selector = Selector::from("Sheet1");
    let at = Reference::new(1, 0).unwrap();
    edit.replace(selector, at, Value::from_comment(original))
        .unwrap();
    let error = edit
        .replace(
            selector,
            at,
            Value::new(original.author(), "Changed protected note").unwrap(),
        )
        .unwrap_err();
    assert!(
        matches!(error,litchi_xls::Error::Refused(ref r) if r.kind()==litchi_core::edit::RefusalKind::Protected)
    );
    let unchanged = edit.commit().unwrap();
    assert!(unchanged.package().to_bytes().unwrap() == bytes);
    assert!(source.to_bytes().unwrap() == bytes);
}

#[test]
fn native_conditional_formats_keep_legacy_links_priorities_and_visual_parameters() {
    use litchi_xls::sheet_features::{Area, FeatureSet, RuleType, Threshold};
    let bytes = native("excel-conditional-drawings.xls");
    let package = Package::from_bytes(&bytes).unwrap();
    let formats = package
        .sheet_features(0, FeatureSet::FORMATS)
        .unwrap()
        .formats;
    assert_eq!(formats.len(), 7);
    assert_eq!(formats.iter().map(|f| f.rule_count).sum::<usize>(), 8);
    for (format, (column, kind, priority, future, count)) in formats.iter().zip([
        (1, RuleType::DataBar, 8, true, 1),
        (2, RuleType::ColorScale, 7, true, 1),
        (3, RuleType::IconSet, 6, true, 1),
        (4, RuleType::Expression, 5, false, 1),
        (5, RuleType::Expression, 4, false, 1),
        (7, RuleType::ColorScale, 3, true, 1),
        (8, RuleType::Expression, 1, false, 2),
    ]) {
        assert_eq!(format.ranges, [Area::new(2, column, 7, column)]);
        assert_eq!(format.future, future);
        assert_eq!(format.rule_count, count);
        let rule = format.rule.as_ref().unwrap();
        assert_eq!(rule.rule_type, kind);
        assert_eq!(rule.priority, priority);
        assert!(!rule.stop_if_true);
    }
    let bar = formats[0].rule.as_ref().unwrap().data_bar.as_ref().unwrap();
    assert_eq!(bar.color, Some(0x638ec6));
    assert_eq!(
        bar.min,
        Threshold {
            kind: 2,
            value: None
        }
    );
    assert_eq!(
        bar.max,
        Threshold {
            kind: 3,
            value: None
        }
    );
    assert_eq!(
        (bar.min_length, bar.max_length, bar.show_value),
        (0, 100, false)
    );
    let scale = formats[1]
        .rule
        .as_ref()
        .unwrap()
        .color_scale
        .as_ref()
        .unwrap();
    assert_eq!(
        scale.points,
        [
            (
                Threshold {
                    kind: 2,
                    value: None
                },
                Some(0xf8696b)
            ),
            (
                Threshold {
                    kind: 5,
                    value: Some("50".into())
                },
                Some(0xffeb84)
            ),
            (
                Threshold {
                    kind: 3,
                    value: None
                },
                Some(0x63be7b)
            ),
        ]
    );
    let icons = formats[2].rule.as_ref().unwrap().icon_set.as_ref().unwrap();
    assert_eq!(
        (icons.icon_set, icons.reverse, icons.show_value),
        (0, false, true)
    );
    assert_eq!(
        icons.thresholds,
        ["0", "33", "67"].map(|v| Threshold {
            kind: 4,
            value: Some(v.into())
        })
    );
    assert_eq!(formats[4].rule.as_ref().unwrap().template, 27); // Duplicate values.

    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let sheet = reader.xls_worksheet(0).unwrap();
    assert_eq!(sheet.conditional_formattings().len(), 3);
    assert_eq!(sheet.conditional_formattings12().len(), 4);
    let links = sheet.conditional_format_extensions();
    assert_eq!(
        links
            .iter()
            .map(|e| (e.identifier(), e.legacy_rule_index(), e.priority()))
            .collect::<Vec<_>>(),
        [
            (4, Some(0), 5),
            (5, Some(0), 4),
            (7, Some(0), 1),
            (7, Some(1), 2)
        ]
    );
    // Formula text can remain opaque in this reader, but the exact native
    // relative-reference expression must remain available to callers.
    assert_eq!(
        sheet.conditional_formattings()[0].rules()[0].formula1_tokens(),
        [
            108, 0, 0, 4, 128, 30, 50, 0, 12, 108, 0, 0, 4, 128, 30, 100, 0, 9, 66, 2, 36, 0
        ]
    );
}

#[test]
fn native_pivot_caption_edit_preserves_cache_streams_items_and_inverse() {
    let bytes = native("excel-visual-analytics.xls");
    let source = Package::from_bytes(&bytes).unwrap();
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let mut expected = reader
        .xls_worksheet(0)
        .unwrap()
        .pivot_table("PivotTable1")
        .unwrap()
        .clone();
    expected.data_items[0].name = "Reviewed revenue".into();
    let mut edit = source.edit_pivot_views().unwrap();
    edit.update_by_name(0, "PivotTable1", |table| {
        table.data_items[0].name = "Reviewed revenue".into()
    })
    .unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    let reopened = Reader::new(Cursor::new(&after)).unwrap();
    assert_eq!(
        reopened
            .xls_worksheet(0)
            .unwrap()
            .pivot_table("PivotTable1"),
        Some(&expected)
    );
    assert_other_streams(&bytes, &after);
    let package = Package::from_bytes(&after).unwrap();
    let undone = package
        .apply_pivot_views(&committed.patch().inverse())
        .unwrap();
    assert!(undone.package().to_bytes().unwrap() == bytes);
    assert!(source.to_bytes().unwrap() == bytes);
}

#[test]
fn native_row_index_keeps_zero_offsets_for_empty_hidden_rows() {
    let bytes = native("excel-date1904-breaks.xls");
    let reader = Reader::new(Cursor::new(&bytes)).unwrap();
    let index = reader
        .xls_worksheet(0)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    let first = &index.blocks()[0];
    assert_eq!(first.dbcell().cell_offsets().len(), 11);
    assert_eq!(&first.dbcell().cell_offsets()[9..], &[0, 0]);
    assert_eq!(first.indexed_rows().len(), 9);
    assert!(
        !first
            .indexed_rows()
            .iter()
            .any(|r| [13, 14].contains(&r.row()))
    );
    assert_eq!(index.first_cell_position(19), Some(17627));
    let stream_map = streams(&bytes);
    let source_dbcell = records(&stream_map[&vec!["Workbook".to_string()]])
        .into_iter()
        .find(|(kind, _)| *kind == 0x00d7)
        .unwrap()
        .1;
    assert_eq!(&first.to_record_bytes().unwrap()[4..], source_dbcell);
}
