use std::io::Cursor;
use std::path::PathBuf;

use litchi_xls::writer::{
    PivotCacheValue, PivotFieldConfig, PivotItemConfig, PivotTableConfig, Writer,
};
use litchi_xls::{
    PivotCacheDateGroupUnit, PivotCacheDateGrouping, PivotCacheDateTime,
    PivotCacheDiscreteGrouping, PivotCacheGrouping, PivotCacheItem, PivotCacheNumericGrouping,
    Workbook,
};

fn item(index: u16) -> PivotItemConfig {
    PivotItemConfig {
        item_type: 0,
        flags: 0,
        cache_index: index,
        name: None,
    }
}

fn config(
    name: &str,
    sheet: &str,
    fields: Vec<PivotFieldConfig>,
    rows: Vec<Vec<PivotCacheValue>>,
) -> PivotTableConfig {
    PivotTableConfig {
        name: name.to_string(),
        source_type: 1,
        source_sheet_name: sheet.to_string(),
        source_first_row: 0,
        source_last_row: rows.len() as u16,
        source_first_col: 0,
        source_last_col: fields
            .iter()
            .filter(|field| !matches!(field.grouping, Some(PivotCacheGrouping::Discrete(_))))
            .count()
            .saturating_sub(1) as u16,
        first_row: 8,
        last_row: 12,
        first_col: 0,
        last_col: 1,
        first_header_row: 8,
        first_data_row: 9,
        first_data_col: 1,
        data_field_name: "Values".to_string(),
        data_axis: 0,
        data_position: 0,
        fields,
        data_items: Vec::new(),
        page_entries: Vec::new(),
        source_data: rows,
    }
}

fn write_configs(configs: Vec<(&str, PivotTableConfig)>) -> Vec<u8> {
    let mut writer = Writer::new();
    for (sheet_name, config) in configs {
        let sheet = writer.add_worksheet(sheet_name).unwrap();
        writer.add_pivot_table(sheet, config).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn numeric_config(sheet: &str) -> PivotTableConfig {
    config(
        "Numeric",
        sheet,
        vec![PivotFieldConfig {
            axis: 1,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: vec![item(0), item(1)],
            name: None,
            cache_name: "Number".into(),
            cache_items: vec![0.0, 5.0, 10.0, 15.0]
                .into_iter()
                .map(PivotCacheItem::Number)
                .collect(),
            is_numeric: true,
            grouping: Some(PivotCacheGrouping::Numeric(PivotCacheNumericGrouping {
                start: 0.0,
                end: 20.0,
                step: 10.0,
                auto_start: true,
                auto_end: false,
                group_items: vec!["0-9".into(), "10-19".into()],
            })),
        }],
        vec![0.0, 5.0, 10.0, 15.0]
            .into_iter()
            .map(|value| vec![PivotCacheValue::Number(value)])
            .collect(),
    )
}

fn date_config(sheet: &str) -> PivotTableConfig {
    let jan = PivotCacheDateTime::try_new(2026, 1, 1, 0, 0, 0).unwrap();
    let feb = PivotCacheDateTime::try_new(2026, 2, 1, 0, 0, 0).unwrap();
    let mar = PivotCacheDateTime::try_new(2026, 3, 1, 0, 0, 0).unwrap();
    config(
        "Dates",
        sheet,
        vec![PivotFieldConfig {
            axis: 1,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: vec![item(0), item(1)],
            name: None,
            cache_name: "Date".into(),
            cache_items: vec![PivotCacheItem::DateTime(jan), PivotCacheItem::DateTime(feb)],
            is_numeric: false,
            grouping: Some(PivotCacheGrouping::Date(PivotCacheDateGrouping {
                unit: PivotCacheDateGroupUnit::Months,
                start: jan,
                end: mar,
                step: 1,
                auto_start: false,
                auto_end: true,
                group_items: vec!["Jan".into(), "Feb".into()],
            })),
        }],
        vec![
            vec![PivotCacheValue::DateTime(jan)],
            vec![PivotCacheValue::DateTime(feb)],
        ],
    )
}

fn discrete_config(sheet: &str) -> PivotTableConfig {
    config(
        "Discrete",
        sheet,
        vec![
            PivotFieldConfig {
                axis: 0,
                subtotal_count: 0,
                subtotal_flags: 0,
                items: vec![item(0), item(1), item(2)],
                name: None,
                cache_name: "Base".into(),
                cache_items: vec!["A".into(), "B".into(), "C".into()],
                is_numeric: false,
                grouping: None,
            },
            PivotFieldConfig {
                axis: 1,
                subtotal_count: 0,
                subtotal_flags: 0,
                items: vec![item(0), item(1)],
                name: None,
                cache_name: "Grouped".into(),
                cache_items: Vec::new(),
                is_numeric: false,
                grouping: Some(PivotCacheGrouping::Discrete(PivotCacheDiscreteGrouping {
                    base_field_index: 0,
                    group_items: vec!["AB".into(), "C".into()],
                    item_to_group: vec![0, 0, 1],
                })),
            },
        ],
        vec![
            vec![PivotCacheValue::StringIndex(0)],
            vec![PivotCacheValue::StringIndex(1)],
            vec![PivotCacheValue::StringIndex(2)],
        ],
    )
}

fn cache_records(bytes: &[u8], id: u16) -> Vec<(u16, Vec<u8>)> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole
        .open_stream(&["_SX_DB_CUR", &format!("{id:04X}")])
        .unwrap();
    let mut records = Vec::new();
    let mut offset = 0;
    while offset < stream.len() {
        let kind = u16::from_le_bytes(stream[offset..offset + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            stream[offset + 2..offset + 4].try_into().unwrap(),
        ));
        records.push((kind, stream[offset + 4..offset + 4 + len].to_vec()));
        offset += 4 + len;
    }
    records
}

fn cache_stream(bytes: &[u8], id: u16) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    ole.open_stream(&["_SX_DB_CUR", &format!("{id:04X}")])
        .unwrap()
}

#[test]
fn emits_and_reopens_exact_numeric_date_and_discrete_grouping_records() {
    let bytes = write_configs(vec![
        ("Numbers", numeric_config("Numbers")),
        ("Dates", date_config("Dates")),
        ("Discrete", discrete_config("Discrete")),
    ]);
    let numeric = cache_records(&bytes, 1);
    let date = cache_records(&bytes, 2);
    let discrete = cache_records(&bytes, 3);
    assert_eq!(
        numeric.iter().find(|record| record.0 == 0x00D8).unwrap().1,
        0x0021u16.to_le_bytes()
    );
    assert_eq!(
        date.iter().find(|record| record.0 == 0x00D8).unwrap().1,
        0x0016u16.to_le_bytes()
    );
    assert_eq!(
        discrete.iter().find(|record| record.0 == 0x00D9).unwrap().1,
        vec![0, 0, 0, 0, 1, 0]
    );
    let sxdb = discrete.iter().find(|record| record.0 == 0x00C6).unwrap();
    assert_eq!(u16::from_le_bytes(sxdb.1[10..12].try_into().unwrap()), 1);
    assert_eq!(u16::from_le_bytes(sxdb.1[12..14].try_into().unwrap()), 2);

    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(
        workbook
            .pivot_caches()
            .iter()
            .map(litchi_xls::PivotCache::stream_id)
            .collect::<Vec<_>>(),
        vec![1, 2, 3]
    );
    assert!(matches!(
        workbook.pivot_caches()[0].fields()[0].grouping(),
        Some(PivotCacheGrouping::Numeric(_))
    ));
    assert!(matches!(
        workbook.pivot_caches()[1].fields()[0].grouping(),
        Some(PivotCacheGrouping::Date(_))
    ));
    assert!(matches!(
        workbook.pivot_caches()[2].fields()[1].grouping(),
        Some(PivotCacheGrouping::Discrete(_))
    ));
}

#[test]
fn reads_libreoffice_number_and_date_grouping_fixtures() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sc/qa/unit/data/xls");
    for file in [
        "pivottable_number_grouping.xls",
        "pivottable_dates_grouping.xls",
    ] {
        let workbook = Workbook::new(std::fs::File::open(root.join(file)).unwrap()).unwrap();
        assert!(
            workbook
                .pivot_caches()
                .iter()
                .flat_map(litchi_xls::PivotCache::fields)
                .any(|field| field.grouping().is_some())
        );
    }
}

#[test]
fn invalid_grouping_is_rejected_without_mutating_writer() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Atomic").unwrap();
    let mut invalid = numeric_config("Atomic");
    if let Some(PivotCacheGrouping::Numeric(grouping)) = &mut invalid.fields[0].grouping {
        grouping.step = 0.0;
    }
    assert!(writer.add_pivot_table(sheet, invalid).is_err());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    assert!(
        Workbook::new(Cursor::new(output.into_inner()))
            .unwrap()
            .pivot_caches()
            .is_empty()
    );

    let mut cyclic = discrete_config("Atomic");
    if let Some(PivotCacheGrouping::Discrete(grouping)) = &mut cyclic.fields[1].grouping {
        grouping.base_field_index = 1;
    }
    assert!(writer.add_pivot_table(sheet, cyclic).is_err());

    let bytes = write_configs(vec![("Numbers", numeric_config("Numbers"))]);
    let mut malformed = cache_stream(&bytes, 1);
    let d8 = malformed
        .windows(4)
        .position(|window| window == [0xD8, 0x00, 0x02, 0x00])
        .unwrap();
    malformed[d8 + 4..d8 + 6].copy_from_slice(&0u16.to_le_bytes());
    assert!(litchi_xls::pivot_table::parse_pivot_cache_stream(&malformed).is_err());
}
