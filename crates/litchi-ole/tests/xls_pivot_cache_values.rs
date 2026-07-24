use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::xls::writer::{
    PivotCacheValue, XlsPivotFieldConfig, XlsPivotItemConfig, XlsPivotTableConfig, XlsWriter,
};
use litchi_ole::xls::{
    PivotCacheDateTime, PivotCacheError, PivotCacheItem, XlsWorkbook,
};

fn typed_items() -> Vec<PivotCacheItem> {
    vec![
        PivotCacheItem::Boolean(false),
        PivotCacheItem::Boolean(true),
        PivotCacheItem::Error(PivotCacheError::DivisionByZero),
        PivotCacheItem::DateTime(PivotCacheDateTime::try_new(2026, 7, 19, 12, 34, 56).unwrap()),
        PivotCacheItem::Empty,
    ]
}

fn typed_config(name: &str, sheet_name: &str) -> XlsPivotTableConfig {
    XlsPivotTableConfig {
        name: name.to_string(),
        source_type: 1,
        source_sheet_name: sheet_name.to_string(),
        source_first_row: 0,
        source_last_row: 5,
        source_first_col: 0,
        source_last_col: 0,
        first_row: 8,
        last_row: 13,
        first_col: 0,
        last_col: 1,
        first_header_row: 8,
        first_data_row: 9,
        first_data_col: 1,
        data_field_name: "Values".to_string(),
        data_axis: 0,
        data_position: 0,
        fields: vec![XlsPivotFieldConfig {
            axis: 1,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: (0..5)
                .map(|index| XlsPivotItemConfig {
                    item_type: 0,
                    flags: 0,
                    cache_index: index,
                    name: None,
                })
                .collect(),
            name: None,
            cache_name: "Mixed".to_string(),
            cache_items: typed_items(),
            is_numeric: false,
            grouping: None,
        }],
        data_items: Vec::new(),
        page_entries: Vec::new(),
        source_data: vec![
            vec![PivotCacheValue::Boolean(false)],
            vec![PivotCacheValue::Boolean(true)],
            vec![PivotCacheValue::Error(PivotCacheError::DivisionByZero)],
            vec![PivotCacheValue::DateTime(
                PivotCacheDateTime::try_new(2026, 7, 19, 12, 34, 56).unwrap(),
            )],
            vec![PivotCacheValue::Empty],
        ],
    }
}

fn write_typed_workbook(sheet_count: usize) -> Vec<u8> {
    let mut writer = XlsWriter::new();
    for index in 0..sheet_count {
        let sheet_name = format!("Typed{}", index + 1);
        let sheet = writer.add_worksheet(&sheet_name).unwrap();
        writer
            .add_pivot_table(sheet, typed_config(&format!("Pivot{}", index + 1), &sheet_name))
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn cache_stream(bytes: &[u8], id: u16) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    ole.open_stream(&["_SX_DB_CUR", &format!("{id:04X}")]).unwrap()
}

fn records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut result = Vec::new();
    let mut offset = 0usize;
    while offset < bytes.len() {
        let kind = u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(bytes[offset + 2..offset + 4].try_into().unwrap()));
        result.push((kind, bytes[offset + 4..offset + 4 + len].to_vec()));
        offset += 4 + len;
    }
    result
}

#[test]
fn typed_cache_items_emit_exact_flags_payloads_and_rows() {
    let bytes = write_typed_workbook(1);
    let cache = cache_stream(&bytes, 1);
    let records = records(&cache);
    let field = records.iter().find(|(kind, _)| *kind == 0x00C7).unwrap();
    assert_eq!(u16::from_le_bytes(field.1[0..2].try_into().unwrap()), 0x0D81);
    assert_eq!(u16::from_le_bytes(field.1[6..8].try_into().unwrap()), 5);
    assert_eq!(u16::from_le_bytes(field.1[12..14].try_into().unwrap()), 5);

    let typed = records
        .iter()
        .filter(|(kind, _)| matches!(*kind, 0x00CA | 0x00CB | 0x00CE | 0x00CF))
        .collect::<Vec<_>>();
    assert_eq!(typed.iter().map(|record| record.0).collect::<Vec<_>>(), vec![0x00CA, 0x00CA, 0x00CB, 0x00CE, 0x00CF]);
    assert_eq!(typed[0].1, 0u16.to_le_bytes());
    assert_eq!(typed[1].1, 1u16.to_le_bytes());
    assert_eq!(typed[2].1, 0x0007u16.to_le_bytes());
    assert_eq!(typed[3].1, vec![0xEA, 0x07, 0x07, 0x00, 19, 12, 34, 56]);
    assert!(typed[4].1.is_empty());
    assert_eq!(
        records.iter().filter(|(kind, _)| *kind == 0x00C8).map(|record| record.1.clone()).collect::<Vec<_>>(),
        vec![vec![0], vec![1], vec![2], vec![3], vec![4]]
    );

    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.pivot_caches().len(), 1);
    assert_eq!(workbook.pivot_caches()[0].fields()[0].items(), typed_items());
    assert_eq!(workbook.pivot_caches()[0].rows().iter().map(|row| row[0].clone()).collect::<Vec<_>>(), typed_items());
}

#[test]
fn typed_caches_coexist_with_workbook_global_stream_ids() {
    let bytes = write_typed_workbook(2);
    assert!(!cache_stream(&bytes, 1).is_empty());
    assert!(!cache_stream(&bytes, 2).is_empty());
    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.pivot_caches().iter().map(|cache| cache.stream_id()).collect::<Vec<_>>(), vec![1, 2]);
    assert_eq!(workbook.pivot_caches()[0].rows(), workbook.pivot_caches()[1].rows());
}

#[test]
fn reads_libreoffice_boolean_and_empty_pivot_caches() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    for (name, expected) in [
        ("pivottable_bool_field_filter.xls", PivotCacheItem::Boolean(true)),
        ("pivottable_empty_item.xls", PivotCacheItem::Empty),
    ] {
        let path = root.join("../../test-data/libreoffice-core/sc/qa/unit/data/xls").join(name);
        let workbook = XlsWorkbook::new(std::fs::File::open(path).unwrap()).unwrap();
        assert!(workbook.pivot_caches().iter().flat_map(|cache| cache.fields()).flat_map(|field| field.items()).any(|item| item == &expected));
    }
}

#[test]
fn malformed_typed_records_are_rejected_and_invalid_add_is_atomic() {
    assert!(PivotCacheError::try_from(1).is_err());
    assert!(PivotCacheDateTime::try_new(2026, 2, 30, 0, 0, 0).is_err());

    let bytes = write_typed_workbook(1);
    let mut cache = cache_stream(&bytes, 1);
    let boolean = cache.windows(4).position(|window| window == [0xCA, 0x00, 0x02, 0x00]).unwrap();
    cache[boolean + 4..boolean + 6].copy_from_slice(&2u16.to_le_bytes());
    assert!(litchi_ole::xls::pivot_table::parse_pivot_cache_stream(&cache).is_err());

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Atomic").unwrap();
    let mut invalid = typed_config("Invalid", "Atomic");
    invalid.fields[0].cache_items.push(PivotCacheItem::Boolean(false));
    assert!(writer.add_pivot_table(sheet, invalid).is_err());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.pivot_caches().is_empty());
}
