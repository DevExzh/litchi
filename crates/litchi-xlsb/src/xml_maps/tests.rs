use super::*;
use crate::package::table::{Column, Range, Table};
use crate::raw::{Records, Writer, kind};
use crate::writer::{MutableWorksheet, WorkbookWriter};
use std::io::Cursor;

fn binding(column_id: u32) -> ColumnBinding {
    ColumnBinding::new(
        column_id,
        7,
        XmlDataType::new(0x2D).unwrap(),
        XPath::new("/root/value").unwrap(),
        true,
    )
    .unwrap()
}

#[test]
fn column_binding_exact_vector() {
    let bytes = serialize_column_binding(&binding(3), Limits::DEFAULT).unwrap();
    let expected = [
        0xDD, 0x02, 0x26, 0x07, 0, 0, 0, 0x02, 0, 0, 0, 0x2D, 0, 0, 0, 0x0B, 0, 0, 0, b'/', 0,
        b'r', 0, b'o', 0, b'o', 0, b't', 0, b'/', 0, b'v', 0, b'a', 0, b'l', 0, b'u', 0, b'e', 0,
        0xDE, 0x02, 0,
    ];
    assert_eq!(bytes, expected);
}

#[test]
fn single_cell_canonical_round_trip_and_envelope() {
    let value = SingleCellBinding::new(
        9,
        1,
        CellReference::new(4, 2).unwrap(),
        7,
        XmlDataType::new(1).unwrap(),
        XPath::new("/r/v").unwrap(),
    )
    .unwrap();
    let bytes = serialize_single_cells(&[value.clone()], Limits::DEFAULT).unwrap();
    let kinds = Records::new(&bytes)
        .map(|record| record.unwrap().kind())
        .collect::<Vec<_>>();
    assert_eq!(kinds.first(), Some(&kind::BEGIN_SINGLE_CELLS));
    assert_eq!(kinds.last(), Some(&kind::END_SINGLE_CELLS));
    let parsed = parse_single_cells(&bytes, Limits::DEFAULT).unwrap();
    assert_eq!(parsed.value(), &[value]);
    assert_eq!(
        patch_single_cells(&parsed, parsed.value(), Limits::DEFAULT).unwrap(),
        bytes
    );
}

#[test]
fn validates_wire_domains_and_absolute_xpath() {
    assert!(XmlDataType::new(0).is_err());
    assert!(XmlDataType::new(0x2E).is_err());
    assert!(XPath::new("relative/path").is_err());
    assert!(XPath::new("/absolute/path").is_ok());
    assert!(XPath::new("/root/item[@a='x']/value").is_ok());
    assert!(XPath::new("/root/item[@a = 'x']/value").is_ok());
    assert!(XPath::new("/ns:root/ns:item/@id").is_ok());
    assert!(XPath::new("/root/child::value").is_err());
    assert!(XPath::new("/root/item[@a='x' and @b='y']/value").is_err());
    assert!(XPath::new("/root//value").is_err());
    assert!(XPath::new("/root/*").is_err());
}

fn rewrite_record(
    source: &[u8],
    target: crate::raw::Kind,
    mut edit: impl FnMut(&mut Vec<u8>),
) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    for record in Records::new(source) {
        let record = record.unwrap();
        let mut payload = record.payload().to_vec();
        if record.kind() == target {
            edit(&mut payload);
        }
        writer.write_record(record.kind(), &payload).unwrap();
    }
    drop(writer);
    output
}

fn minimal_table(binding: &ColumnBinding, ignored_flags: u32) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Writer::new(&mut output);
    let mut list = vec![0; 24];
    list[16..20].copy_from_slice(&2u32.to_le_bytes());
    list[20..24].copy_from_slice(&2u32.to_le_bytes());
    writer.write_record(kind::BEGIN_LIST, &list).unwrap();
    writer
        .write_record(kind::BEGIN_LIST_COLS, &1u32.to_le_bytes())
        .unwrap();
    writer
        .write_record(kind::BEGIN_LIST_COL, &binding.column_id().to_le_bytes())
        .unwrap();
    let encoded = serialize_column_binding(binding, Limits::DEFAULT).unwrap();
    for record in Records::new(&encoded) {
        let record = record.unwrap();
        let mut payload = record.payload().to_vec();
        if record.kind() == kind::BEGIN_LIST_XML_CPR {
            let flags = u32::from_le_bytes(payload[4..8].try_into().unwrap()) | ignored_flags;
            payload[4..8].copy_from_slice(&flags.to_le_bytes());
        }
        writer.write_record(record.kind(), &payload).unwrap();
    }
    writer.write_record(kind::END_LIST_COL, &[]).unwrap();
    writer.write_record(kind::END_LIST_COLS, &[]).unwrap();
    writer.write_record(kind::END_LIST, &[]).unwrap();
    drop(writer);
    output
}

#[test]
fn ordinary_table_patch_preserves_ignored_xml_property_flags() {
    let source_bytes = minimal_table(&binding(3), 0x8000_0001);
    let source = parse_table_bindings(&source_bytes, Limits::DEFAULT).unwrap();
    let changed = MappedTable::new(
        2,
        vec![
            ColumnBinding::new(
                3,
                7,
                XmlDataType::new(0x2D).unwrap(),
                XPath::new("/root/changed").unwrap(),
                false,
            )
            .unwrap(),
        ],
    )
    .unwrap();
    let patched = patch_table_bindings(&source, &changed, Limits::DEFAULT).unwrap();
    let xml = Records::new(&patched)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == kind::BEGIN_LIST_XML_CPR)
        .unwrap();
    assert_eq!(
        u32::from_le_bytes(xml.payload()[4..8].try_into().unwrap()),
        0x8000_0001
    );
}

#[test]
fn single_cell_accepts_but_will_not_normalize_unmodeled_metadata() {
    let value = SingleCellBinding::new(
        9,
        1,
        CellReference::new(4, 2).unwrap(),
        7,
        XmlDataType::new(1).unwrap(),
        XPath::new("/r/v").unwrap(),
    )
    .unwrap();
    let canonical = serialize_single_cells(&[value.clone()], Limits::DEFAULT).unwrap();
    let with_list_metadata = rewrite_record(&canonical, kind::BEGIN_LIST, |payload| {
        payload[32..36].copy_from_slice(&0x1Fu32.to_le_bytes());
        payload[60..64].copy_from_slice(&17u32.to_le_bytes());
    });
    let with_column_metadata =
        rewrite_record(&with_list_metadata, kind::BEGIN_LIST_COL, |payload| {
            payload[4..8].copy_from_slice(&6u32.to_le_bytes());
            let mut total = Vec::new();
            Writer::new(&mut total).write_wide_string("Total").unwrap();
            payload.splice(32..36, total);
        });
    let source_bytes = rewrite_record(&with_column_metadata, kind::BEGIN_LIST_XML_CPR, |payload| {
        payload[4..8].copy_from_slice(&0x8000_0003u32.to_le_bytes());
    });
    let source = parse_single_cells(&source_bytes, Limits::DEFAULT).unwrap();
    assert_eq!(source.value(), &[value.clone()]);
    assert_eq!(source.connection_ids(), &[17]);
    assert_eq!(
        patch_single_cells(&source, source.value(), Limits::DEFAULT).unwrap(),
        source_bytes
    );
    let changed = SingleCellBinding::new(
        9,
        1,
        value.cell(),
        7,
        value.data_type(),
        XPath::new("/r/changed").unwrap(),
    )
    .unwrap();
    assert!(patch_single_cells(&source, &[changed], Limits::DEFAULT).is_err());
}

#[test]
fn single_cell_preserves_ignored_list_flags_and_rejects_unknown_total_function() {
    let value = SingleCellBinding::new(
        9,
        1,
        CellReference::new(4, 2).unwrap(),
        7,
        XmlDataType::new(1).unwrap(),
        XPath::new("/r/v").unwrap(),
    )
    .unwrap();
    let canonical = serialize_single_cells(&[value], Limits::DEFAULT).unwrap();
    let ignored = rewrite_record(&canonical, kind::BEGIN_LIST, |payload| {
        payload[32..36].copy_from_slice(&(2u32 | 0x20).to_le_bytes());
    });
    let source = parse_single_cells(&ignored, Limits::DEFAULT).unwrap();
    assert_eq!(
        patch_single_cells(&source, source.value(), Limits::DEFAULT).unwrap(),
        ignored
    );
    let changed = SingleCellBinding::new(
        9,
        1,
        source.value()[0].cell(),
        7,
        source.value()[0].data_type(),
        XPath::new("/r/changed").unwrap(),
    )
    .unwrap();
    assert!(patch_single_cells(&source, &[changed], Limits::DEFAULT).is_err());
    let unknown_total = rewrite_record(&canonical, kind::BEGIN_LIST_COL, |payload| {
        payload[4..8].copy_from_slice(&10u32.to_le_bytes());
    });
    assert!(parse_single_cells(&unknown_total, Limits::DEFAULT).is_err());
    assert!(
        SingleCellBinding::new(
            u32::MAX,
            1,
            CellReference::new(4, 2).unwrap(),
            7,
            XmlDataType::new(1).unwrap(),
            XPath::new("/r/v").unwrap(),
        )
        .is_err()
    );
}

#[test]
fn opaque_record_limits_accept_exact_boundaries_and_reject_excess() {
    let value = SingleCellBinding::new(
        9,
        1,
        CellReference::new(4, 2).unwrap(),
        7,
        XmlDataType::new(1).unwrap(),
        XPath::new("/r/v").unwrap(),
    )
    .unwrap();
    let canonical = serialize_single_cells(&[value], Limits::DEFAULT).unwrap();
    let mut source = Vec::new();
    let mut writer = Writer::new(&mut source);
    for record in Records::new(&canonical) {
        let record = record.unwrap();
        writer
            .write_record(record.kind(), record.payload())
            .unwrap();
        if record.kind() == kind::BEGIN_SINGLE_CELLS {
            writer
                .write_record(crate::raw::Kind::new(3_000).unwrap(), &[1, 2, 3])
                .unwrap();
        }
    }
    drop(writer);
    let exact = Limits {
        max_opaque_records: 1,
        max_opaque_bytes: 3,
        ..Limits::DEFAULT
    };
    assert!(parse_single_cells(&source, exact).is_ok());
    assert!(
        parse_single_cells(
            &source,
            Limits {
                max_opaque_records: 0,
                ..exact
            }
        )
        .is_err()
    );
    assert!(
        parse_single_cells(
            &source,
            Limits {
                max_opaque_bytes: 2,
                ..exact
            }
        )
        .is_err()
    );

    let table = minimal_table(&binding(3), 0);
    let mut table_source = Vec::new();
    let mut writer = Writer::new(&mut table_source);
    for record in Records::new(&table) {
        let record = record.unwrap();
        writer
            .write_record(record.kind(), record.payload())
            .unwrap();
        if record.kind() == kind::BEGIN_LIST {
            writer
                .write_record(crate::raw::Kind::new(3_000).unwrap(), &[1, 2, 3])
                .unwrap();
        }
    }
    drop(writer);
    assert!(parse_table_bindings(&table_source, exact).is_ok());
    assert!(
        parse_table_bindings(
            &table_source,
            Limits {
                max_opaque_records: 0,
                ..exact
            }
        )
        .is_err()
    );
}

#[test]
fn ordinary_unmapped_table_does_not_require_map_info() {
    let mut worksheet = MutableWorksheet::new("Plain");
    worksheet.set_cell(0, 0, "Value");
    worksheet.set_cell(1, 0, "data");
    worksheet
        .add_table(Table {
            id: 17,
            display_name: Some("PlainTable".to_string()),
            range: Range {
                first_row: 0,
                last_row: 1,
                first_column: 0,
                last_column: 0,
            },
            header_row_count: 1,
            columns: vec![Column {
                id: 1,
                name: Some("Value".to_string()),
                ..Column::default()
            }],
            ..Table::default()
        })
        .unwrap();
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(worksheet);
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).unwrap();

    let workbook = crate::Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let snapshot = workbook.xml_maps().unwrap();
    assert!(snapshot.map_info().is_none());
    assert!(snapshot.mapped_tables().is_empty());
    assert!(
        snapshot
            .source()
            .dependencies()
            .iter()
            .any(|part| part.content_type() == "application/vnd.ms-excel.table")
    );
}
