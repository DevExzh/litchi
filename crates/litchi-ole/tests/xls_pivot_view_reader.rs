use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{
    PivotCacheValue, XlsPivotDataItemConfig, XlsPivotFieldConfig, XlsPivotItemConfig,
    XlsPivotTableConfig, XlsWriter,
};
use litchi_ole::xls::{
    PivotAxis, PivotAxisField, PivotCacheGrouping, PivotCacheItem, PivotCacheNumericGrouping,
    XlsPivotViewEditor, XlsWorkbook,
};

fn generated_workbook() -> Vec<u8> {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Source").unwrap();
    writer.add_pivot_table(sheet, XlsPivotTableConfig {
        name: "SalesPivot".to_string(), source_type: 1, source_sheet_name: "Source".to_string(),
        source_first_row: 0, source_last_row: 2, source_first_col: 0, source_last_col: 1,
        first_row: 8, last_row: 10, first_col: 0, last_col: 1,
        first_header_row: 8, first_data_row: 9, first_data_col: 1,
        data_field_name: "Values".to_string(), data_axis: 0, data_position: 0,
        fields: vec![
            XlsPivotFieldConfig {
                axis: 1, subtotal_count: 0, subtotal_flags: 0,
                items: vec![
                    XlsPivotItemConfig { item_type: 0, flags: 0, cache_index: 0, name: None },
                    XlsPivotItemConfig { item_type: 0, flags: 0, cache_index: 1, name: None },
                ],
                name: None, cache_name: "Region".to_string(),
                cache_items: vec![PivotCacheItem::from("East"), PivotCacheItem::from("West")],
                is_numeric: false, grouping: None,
            },
            XlsPivotFieldConfig {
                axis: 8, subtotal_count: 0, subtotal_flags: 0, items: Vec::new(), name: None,
                cache_name: "Amount".to_string(), cache_items: Vec::new(), is_numeric: true,
                grouping: None,
            },
        ],
        data_items: vec![XlsPivotDataItemConfig {
            source_field_index: 1, function: 0, display_format: 0,
            base_field_index: 0, base_item_index: 0, num_format_index: 0,
            name: "Sum of Amount".to_string(),
        }],
        page_entries: Vec::new(),
        source_data: vec![
            vec![PivotCacheValue::StringIndex(0), PivotCacheValue::Number(10.0)],
            vec![PivotCacheValue::StringIndex(1), PivotCacheValue::Number(20.0)],
        ],
    }).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn writer_view_records_reopen_with_identical_layout_and_cache_link() {
    let workbook = XlsWorkbook::new(Cursor::new(generated_workbook())).unwrap();
    let table = workbook.xls_worksheet(0).unwrap().pivot_table("SalesPivot").unwrap();
    assert_eq!(table.cache_index(), 0);
    assert_eq!(table.view.data_row_count, 2);
    assert_eq!(table.view.data_col_count, 1);
    assert_eq!(table.fields.len(), 2);
    assert_eq!(table.fields[0].axis, PivotAxis::Row);
    assert_eq!(table.fields[0].items.len(), 2);
    assert_eq!(table.fields[0].extension.as_ref().unwrap().flags, 0x0AA0_141E);
    assert_eq!(table.row_fields, [PivotAxisField::Field(0)]);
    assert!(table.column_fields.is_empty());
    assert_eq!(table.row_lines.len(), 2);
    assert_eq!(table.column_lines.len(), 1);
    assert_eq!(table.extension.as_ref().unwrap().flags, 0x004F_0200);
    assert_eq!(table.query_tag.as_ref().unwrap().table_name, "SalesPivot");
    assert_eq!(table.view_ex9.as_ref().unwrap().auto_format_index, 1);
    assert!(table.additional_extensions.iter().any(|extension| extension.payload == [0x08, 0x41, 0x40, 0, 0, 0]));
    assert!(table.fields.iter().all(|field| !field.additional_extensions.is_empty()));

    let cache = workbook.pivot_cache_for_table(table).unwrap();
    assert_eq!(cache.stream_id(), 1);
    assert_eq!(table.cache_field(cache, 0).unwrap().name(), "Region");
}

#[test]
fn libreoffice_pivot_views_parse_with_cache_links() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sc/qa/unit/data/xls");
    for name in ["pivottable_number_grouping.xls", "pivottable_dates_grouping.xls"] {
        let workbook = XlsWorkbook::new(std::fs::File::open(root.join(name)).unwrap()).unwrap();
        let tables = (0..workbook.sheets().len())
            .filter_map(|index| workbook.worksheet_pivot_tables(index).ok())
            .flatten()
            .collect::<Vec<_>>();
        assert!(!tables.is_empty(), "{name}");
        for table in tables { assert!(workbook.pivot_cache_for_table(table).is_ok(), "{name}"); }
    }
}

#[test]
fn pivot_view_editor_is_byte_exact_for_noop_and_reopens_mutations() {
    let original = generated_workbook();
    assert_eq!(XlsPivotViewEditor::new(original.clone()).unwrap().finish().unwrap(), original);

    let mut editor = XlsPivotViewEditor::new(original).unwrap();
    editor.update_by_name(0, "SalesPivot", |table| {
        table.view.name = "RenamedPivot".to_string();
        table.view.first_row += 2;
        table.view.last_row += 2;
        table.view.first_header_row += 2;
        table.view.first_data_row += 2;
    }).unwrap();
    let rewritten = editor.finish().unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(rewritten)).unwrap();
    let table = workbook.xls_worksheet(0).unwrap().pivot_table("RenamedPivot").unwrap();
    assert_eq!(table.view.first_row, 10);
    assert_eq!(table.query_tag.as_ref().unwrap().table_name, "RenamedPivot");
    assert_eq!(workbook.pivot_cache_for_table(table).unwrap().stream_id(), 1);
}

#[test]
fn pivot_view_editor_rolls_back_invalid_cache_and_supports_remove_add() {
    let original = generated_workbook();
    let mut editor = XlsPivotViewEditor::new(original.clone()).unwrap();
    assert!(editor.reassign_cache_by_name(0, "SalesPivot", 99).is_err());
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = XlsPivotViewEditor::new(original).unwrap();
    let table = editor.remove_by_name(0, "SalesPivot").unwrap();
    editor.add(0, table).unwrap();
    let reopened = XlsWorkbook::new(Cursor::new(editor.finish().unwrap())).unwrap();
    assert!(reopened.xls_worksheet(0).unwrap().pivot_table("SalesPivot").is_some());
}

#[test]
fn pivot_view_editor_regenerates_grouping_cache_and_preserves_fixture_noops() {
    let mut editor = XlsPivotViewEditor::new(generated_workbook()).unwrap();
    editor.update_cache_grouping(0, 1, Some(PivotCacheGrouping::Numeric(PivotCacheNumericGrouping {
        start: 0.0, end: 30.0, step: 10.0, auto_start: false, auto_end: false,
        group_items: vec!["0-9".into(), "10-19".into(), "20-29".into()],
    }))).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(editor.finish().unwrap())).unwrap();
    assert!(matches!(workbook.pivot_caches()[0].fields()[1].grouping(), Some(PivotCacheGrouping::Numeric(_))));

    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/libreoffice-core/sc/qa/unit/data/xls/pivottable_number_grouping.xls");
    let bytes = std::fs::read(fixture).unwrap();
    assert_eq!(XlsPivotViewEditor::new(bytes.clone()).unwrap().finish().unwrap(), bytes);
}
