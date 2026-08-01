//! Unit tests for table authoring and Escher emission.

use super::*;
use crate::escher::parser::EscherParser;
use crate::escher::record::EscherRecord;
use crate::escher::shape::{EscherShape, EscherShapeType};
use crate::escher::types::EscherRecordType;

fn positioned(table: Table, x: i32, y: i32) -> PositionedTable {
    PositionedTable { x, y, table }
}

fn parse_emitted_group(bytes: &[u8]) -> EscherShape<'_> {
    let parser = EscherParser::new(bytes);
    let root = parser
        .root_container()
        .expect("root container")
        .expect("valid container");
    assert_eq!(root.record().record_type, EscherRecordType::SpgrContainer);
    EscherShape::from_container(root)
}

#[test]
fn new_rejects_invalid_dimensions() {
    assert!(Table::new(0, 3).is_err());
    assert!(Table::new(2, 0).is_err());
    assert!(Table::new(MAX_TABLE_DIMENSION + 1, 1).is_err());
    assert!(Table::new(1, MAX_TABLE_DIMENSION + 1).is_err());
    assert!(Table::new(1, 1).is_ok());
}

#[test]
fn cell_text_access_is_bounds_checked() {
    let mut table = Table::new(2, 3).unwrap();
    assert_eq!(table.rows(), 2);
    assert_eq!(table.columns(), 3);
    assert_eq!(table.cell(0, 0), Some(""));
    assert_eq!(table.cell(2, 0), None);
    assert_eq!(table.cell(0, 3), None);

    table.set_cell_text(1, 2, "last").unwrap();
    assert_eq!(table.cell(1, 2), Some("last"));
    assert!(table.set_cell_text(2, 0, "x").is_err());
    assert!(table.set_cell_text(0, 3, "x").is_err());
}

#[test]
fn dimension_setters_validate_and_convert() {
    let mut table = Table::new(2, 2).unwrap();
    assert_eq!(table.column_width(0), Some(DEFAULT_COLUMN_WIDTH_PT));
    assert_eq!(table.row_height(1), Some(DEFAULT_ROW_HEIGHT_PT));
    assert_eq!(table.column_width(2), None);
    assert_eq!(table.row_height(2), None);

    table.set_column_width(1, 120).unwrap();
    table.set_row_height(0, 50).unwrap();
    assert_eq!(table.column_width(1), Some(120));
    assert_eq!(table.row_height(0), Some(50));

    assert!(table.set_column_width(2, 10).is_err());
    assert!(table.set_column_width(0, 0).is_err());
    assert!(table.set_row_height(2, 10).is_err());
    assert!(table.set_row_height(0, -5).is_err());
}

#[test]
fn shape_count_covers_group_and_cells() {
    assert_eq!(Table::new(2, 3).unwrap().shape_count(), 7);
    assert_eq!(Table::new(1, 1).unwrap().shape_count(), 2);
}

#[test]
fn emitted_group_is_detected_as_table_with_grid_and_text() {
    let mut table = Table::new(2, 3).unwrap();
    let texts = ["A1", "B1", "C1", "A2", "B2", "C2"];
    for (index, text) in texts.iter().enumerate() {
        table.set_cell_text(index / 3, index % 3, *text).unwrap();
    }

    let bytes =
        build_table_spgr_container(&positioned(table, pt_to_emu_i32(50), pt_to_emu_i32(60)), 5)
            .unwrap();
    let shape = parse_emitted_group(&bytes);

    assert_eq!(shape.shape_type(), EscherShapeType::Table);
    assert_eq!(shape.children().len(), 6);

    // Grid: two distinct row tops, three distinct column lefts.
    let lefts: std::collections::BTreeSet<i32> = shape
        .children()
        .iter()
        .map(|cell| cell.anchor().unwrap().left)
        .collect();
    let tops: std::collections::BTreeSet<i32> = shape
        .children()
        .iter()
        .map(|cell| cell.anchor().unwrap().top)
        .collect();
    assert_eq!(lefts.len(), 3);
    assert_eq!(tops.len(), 2);

    for (index, text) in texts.iter().enumerate() {
        assert_eq!(shape.children()[index].text().as_deref(), Some(*text));
    }

    // Default cell: 100x40 pt -> 800x320 master units.
    let first = shape.children()[0].anchor().unwrap();
    assert_eq!((first.left, first.top), (0, 0));
    assert_eq!((first.width(), first.height()), (800, 320));

    // Group anchor: (50pt, 60pt) -> (400, 480) master units. ClientAnchor is
    // host-defined, so validate it through the typed PPT projection.
    let (record, consumed) = litchi_odraw::Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    let typed = litchi_odraw::shape::Shape::try_from(record).unwrap();
    assert_eq!(typed.client_anchor().unwrap().len(), 8);
    let group_anchor = crate::ppt::odraw::anchor(&typed).unwrap().unwrap();
    assert_eq!((group_anchor.left(), group_anchor.top()), (400, 480));
    assert_eq!((group_anchor.width(), group_anchor.height()), (2400, 640));
}

#[test]
fn custom_cell_dimensions_drive_cell_anchors() {
    let mut table = Table::new(1, 2).unwrap();
    table.set_column_width(0, 120).unwrap();
    table.set_column_width(1, 80).unwrap();
    table.set_row_height(0, 50).unwrap();

    let bytes = build_table_spgr_container(&positioned(table, 0, 0), 3).unwrap();
    let shape = parse_emitted_group(&bytes);

    let first = shape.children()[0].anchor().unwrap();
    let second = shape.children()[1].anchor().unwrap();
    // 120pt -> 960 master units, 80pt -> 640, 50pt -> 400.
    assert_eq!((first.left, first.width(), first.height()), (0, 960, 400));
    assert_eq!((second.left, second.width()), (960, 640));
}

#[test]
fn tertiary_opt_carries_table_flag_and_row_heights() {
    let mut table = Table::new(2, 1).unwrap();
    table.set_row_height(1, 60).unwrap();

    let bytes = build_table_spgr_container(&positioned(table, 0, 0), 3).unwrap();
    let parser = EscherParser::new(&bytes);
    let root = parser.root_container().unwrap().unwrap();
    let header = root.find_child(EscherRecordType::SpContainer).unwrap();
    let header = crate::escher::container::EscherContainer::new(header);
    let opt = header.find_child(EscherRecordType::TertiaryOpt).unwrap();

    // Two properties: table flag + complex row-height array.
    assert_eq!(opt.instance, 2);
    assert_eq!(opt.data.len(), 2 * 6 + 6 + 2 * 4);

    let prop0 = u16::from_le_bytes([opt.data[0], opt.data[1]]);
    let value0 = u32::from_le_bytes(opt.data[2..6].try_into().unwrap());
    assert_eq!((prop0, value0), (PROP_GROUP_TABLE_PROPERTIES, 1));

    let prop1 = u16::from_le_bytes([opt.data[6], opt.data[7]]);
    let length1 = u32::from_le_bytes(opt.data[8..12].try_into().unwrap());
    assert_eq!(
        prop1,
        PROP_GROUP_TABLE_ROW_PROPERTIES | PROPERTY_FLAG_COMPLEX
    );
    assert_eq!(length1, 6 + 2 * 4);

    // Complex array header: nElems, nElemsAlloc, cbElem.
    let complex = &opt.data[12..];
    assert_eq!(u16::from_le_bytes([complex[0], complex[1]]), 2);
    assert_eq!(u16::from_le_bytes([complex[2], complex[3]]), 2);
    assert_eq!(u16::from_le_bytes([complex[4], complex[5]]), 4);
    // Row heights: 40pt -> 320, 60pt -> 480 master units.
    assert_eq!(i32::from_le_bytes(complex[6..10].try_into().unwrap()), 320);
    assert_eq!(i32::from_le_bytes(complex[10..14].try_into().unwrap()), 480);
}

#[test]
fn header_records_appear_in_poi_order() {
    let table = Table::new(1, 1).unwrap();
    let bytes = build_table_spgr_container(&positioned(table, 0, 0), 9).unwrap();
    let parser = EscherParser::new(&bytes);
    let root = parser.root_container().unwrap().unwrap();
    let header = root.find_child(EscherRecordType::SpContainer).unwrap();
    let header = crate::escher::container::EscherContainer::new(header);

    let order: Vec<EscherRecordType> = header
        .children()
        .flatten()
        .map(|record: EscherRecord<'_>| record.record_type)
        .collect();
    assert_eq!(
        order,
        vec![
            EscherRecordType::Spgr,
            EscherRecordType::Sp,
            EscherRecordType::TertiaryOpt,
            EscherRecordType::ClientAnchor,
        ]
    );
}

#[test]
fn oversized_table_dimensions_return_an_error() {
    let mut table = Table::new(1, 2).unwrap();
    table.set_column_width(0, 160_000).unwrap();
    table.set_column_width(1, 160_000).unwrap();

    let error = build_table_spgr_container(&positioned(table, 0, 0), 3)
        .expect_err("the aggregate width exceeds i32 EMUs");
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
}

#[test]
fn cell_shape_ids_are_consecutive_after_group_id() {
    let table = Table::new(2, 2).unwrap();
    let bytes = build_table_spgr_container(&positioned(table, 0, 0), 11).unwrap();
    let shape = parse_emitted_group(&bytes);

    let ids: Vec<u32> = shape
        .children()
        .iter()
        .filter_map(|c| c.shape_id())
        .collect();
    assert_eq!(ids, vec![12, 13, 14, 15]);
}

#[test]
fn empty_cell_text_still_emits_a_textbox() {
    let table = Table::new(1, 1).unwrap();
    let bytes = build_table_spgr_container(&positioned(table, 0, 0), 3).unwrap();
    let shape = parse_emitted_group(&bytes);

    assert_eq!(shape.children().len(), 1);
    // Textbox exists but holds no characters.
    assert!(shape.children()[0].text().is_none());
}

#[test]
fn dg_container_accounts_for_table_shapes() {
    use super::super::escher::create_dg_container_with_tables;

    let mut table = Table::new(2, 3).unwrap();
    table.set_cell_text(0, 0, "A1").unwrap();
    let tables = [positioned(table, 0, 0)];

    let bytes = create_dg_container_with_tables(2, &[], &tables).unwrap();
    let parser = EscherParser::new(&bytes);
    let root = parser.root_container().unwrap().unwrap();
    let dg = root.find_child(EscherRecordType::Dg).unwrap();

    // csp = group + background + table group + 6 cells = 9.
    let csp = u32::from_le_bytes(dg.data[0..4].try_into().unwrap());
    assert_eq!(csp, 9);
    // spidCur = (drawing_id << 10) + csp.
    let spid_cur = u32::from_le_bytes(dg.data[4..8].try_into().unwrap());
    assert_eq!(spid_cur, (2 << 10) + 9);

    // The slide-level group contains the patriarch and the table group.
    let spgr = root.find_child(EscherRecordType::SpgrContainer).unwrap();
    let spgr = crate::escher::container::EscherContainer::new(spgr);
    let table_groups = spgr.find_children(EscherRecordType::SpgrContainer);
    assert_eq!(table_groups.len(), 1);
}
