use std::io::Cursor;

use litchi_xls::XlsWorkbook;
use litchi_xls::writer::{
    PivotCacheValue, XlsPivotDataItemConfig, XlsPivotFieldConfig, XlsPivotTableConfig, XlsWriter,
};

const SXVIEW: u16 = 0x00B0;
const SXSTREAMID: u16 = 0x00D5;
const DCONREF: u16 = 0x0051;

fn pivot_config(
    name: &str,
    source_sheet_name: &str,
    first_row: u16,
    value: f64,
) -> XlsPivotTableConfig {
    XlsPivotTableConfig {
        name: name.to_string(),
        source_type: 1,
        source_sheet_name: source_sheet_name.to_string(),
        source_first_row: 0,
        source_last_row: 1,
        source_first_col: 0,
        source_last_col: 0,
        first_row,
        last_row: first_row + 1,
        first_col: 0,
        last_col: 1,
        first_header_row: first_row,
        first_data_row: first_row + 1,
        first_data_col: 1,
        data_field_name: "Values".to_string(),
        data_axis: 2,
        data_position: 0,
        fields: vec![XlsPivotFieldConfig {
            axis: 8,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: Vec::new(),
            name: None,
            cache_name: "Value".to_string(),
            cache_items: Vec::new(),
            is_numeric: true,
            grouping: None,
        }],
        data_items: vec![XlsPivotDataItemConfig {
            source_field_index: 0,
            function: 0,
            display_format: 0,
            base_field_index: 0,
            base_item_index: 0,
            num_format_index: 0,
            name: "Sum of Value".to_string(),
        }],
        page_entries: Vec::new(),
        source_data: vec![vec![PivotCacheValue::Number(value)]],
    }
}

fn generated_bytes(pivots_per_sheet: &[usize]) -> Vec<u8> {
    let mut writer = XlsWriter::new();
    for (sheet_index, &pivot_count) in pivots_per_sheet.iter().enumerate() {
        let sheet_name = format!("Source{}", sheet_index + 1);
        let sheet = writer.add_worksheet(&sheet_name).unwrap();
        for pivot_index in 0..pivot_count {
            let ordinal = pivots_per_sheet[..sheet_index].iter().sum::<usize>() + pivot_index;
            writer
                .add_pivot_table(
                    sheet,
                    pivot_config(
                        &format!("Pivot{}", ordinal + 1),
                        &sheet_name,
                        4 + u16::try_from(pivot_index).unwrap() * 4,
                        11.0 + ordinal as f64 * 11.0,
                    ),
                )
                .unwrap();
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn workbook_records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes(stream[offset..offset + 2].try_into().unwrap());
        let length = usize::from(u16::from_le_bytes(
            stream[offset + 2..offset + 4].try_into().unwrap(),
        ));
        let end = offset + 4 + length;
        records.push((record_type, stream[offset + 4..end].to_vec()));
        offset = end;
    }
    records
}

fn record_u16_values(records: &[(u16, Vec<u8>)], record_type: u16, offset: usize) -> Vec<u16> {
    records
        .iter()
        .filter(|(kind, data)| *kind == record_type && data.len() >= offset + 2)
        .map(|(_, data)| u16::from_le_bytes(data[offset..offset + 2].try_into().unwrap()))
        .collect()
}

fn cache_streams(bytes: &[u8], count: u16) -> Vec<Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    (1..=count)
        .map(|stream_id| {
            let name = format!("{stream_id:04X}");
            ole.open_stream(&["_SX_DB_CUR", &name]).unwrap()
        })
        .collect()
}

fn contains_f64(bytes: &[u8], value: f64) -> bool {
    let encoded = value.to_le_bytes();
    bytes.windows(encoded.len()).any(|window| window == encoded)
}

#[test]
fn pivot_cache_links_progress_across_two_worksheets() {
    let bytes = generated_bytes(&[1, 1]);
    let records = workbook_records(&bytes);
    assert_eq!(record_u16_values(&records, SXVIEW, 14), vec![0, 1]);
    assert_eq!(record_u16_values(&records, SXSTREAMID, 0), vec![1, 2]);

    let dconrefs = records
        .iter()
        .filter(|(record_type, _)| *record_type == DCONREF)
        .map(|(_, data)| data.as_slice())
        .collect::<Vec<_>>();
    assert_eq!(dconrefs.len(), 2);
    assert!(dconrefs[0].windows(7).any(|window| window == b"Source1"));
    assert!(dconrefs[1].windows(7).any(|window| window == b"Source2"));

    let caches = cache_streams(&bytes, 2);
    assert_ne!(caches[0], caches[1]);
    assert!(contains_f64(&caches[0], 11.0));
    assert!(!contains_f64(&caches[0], 22.0));
    assert!(contains_f64(&caches[1], 22.0));
    assert!(!contains_f64(&caches[1], 11.0));

    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.sheets().len(), 2);
    assert!(workbook.xls_worksheet(0).is_ok());
    assert!(workbook.xls_worksheet(1).is_ok());
}

#[test]
fn pivot_cache_links_progress_from_same_sheet_to_later_sheet() {
    let bytes = generated_bytes(&[2, 1]);
    let records = workbook_records(&bytes);
    assert_eq!(record_u16_values(&records, SXVIEW, 14), vec![0, 1, 2]);
    assert_eq!(record_u16_values(&records, SXSTREAMID, 0), vec![1, 2, 3]);

    let caches = cache_streams(&bytes, 3);
    for (cache_index, expected_value) in [11.0, 22.0, 33.0].into_iter().enumerate() {
        assert!(contains_f64(&caches[cache_index], expected_value));
    }

    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(workbook.sheets().len(), 2);
    assert!(workbook.xls_worksheet(0).is_ok());
    assert!(workbook.xls_worksheet(1).is_ok());
}
