//! Regression coverage for BIFF8 worksheet row-block indexes.

use super::codec::RowBlockIndexCollector;
use super::model::{DbCellRecord, WorksheetIndexRecord};
use super::{
    BOF_RECORD_TYPE, DBCELL_RECORD_TYPE, DEF_COL_WIDTH_RECORD_TYPE, EOF_RECORD_TYPE,
    INDEX_FIXED_LEN, INDEX_RECORD_TYPE, MAX_ROW_BLOCKS, ROW_RECORD_TYPE,
};

const SIMPLE_INDEX: [u8; 20] = [
    0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0x66, 0x06, 0, 0, 0xA0, 0x06, 0, 0,
];
const SIMPLE_DBCELL: [u8; 6] = [0x22, 0, 0, 0, 0, 0];
const SIMPLE_ROW: [u8; 16] = [0, 0, 0, 0, 1, 0, 0xFF, 0, 0, 0, 0, 0, 0, 1, 0x0F, 0];
const SIMPLE_LABEL_SST: [u8; 10] = [0, 0, 0, 0, 0x0F, 0, 0, 0, 0, 0];

fn simple_collector() -> RowBlockIndexCollector {
    let mut collector = RowBlockIndexCollector::new(2_500, 1_450);
    collector.feed_record(1_450, BOF_RECORD_TYPE, &[0; 16]);
    collector.feed_record(1_470, INDEX_RECORD_TYPE, &SIMPLE_INDEX);
    collector.feed_record(1_638, DEF_COL_WIDTH_RECORD_TYPE, &[8, 0]);
    collector.feed_record(1_662, ROW_RECORD_TYPE, &SIMPLE_ROW);
    collector.feed_record(1_682, 0x00FD, &SIMPLE_LABEL_SST);
    collector.feed_record(1_696, DBCELL_RECORD_TYPE, &SIMPLE_DBCELL);
    collector.feed_record(1_750, EOF_RECORD_TYPE, &[]);
    collector
}

#[test]
fn parses_resolves_and_round_trips_poi_simple_reference() {
    let index = WorksheetIndexRecord::parse_payload(&SIMPLE_INDEX, 2_500).unwrap();
    assert_eq!(
        (index.first_data_row(), index.last_data_row_exclusive()),
        (0, 1)
    );
    assert_eq!(index.default_column_width_position(), 1_638);
    assert_eq!(index.dbcell_positions(), &[1_696]);
    let dbcell = DbCellRecord::parse_payload(1_696, 2_500, &SIMPLE_DBCELL).unwrap();
    assert_eq!(dbcell.first_row_position(), Some(1_662));
    assert_eq!(dbcell.resolve_cell_positions(1_682).unwrap(), vec![1_682]);

    let aggregate = simple_collector().finish().unwrap().unwrap();
    assert_eq!(aggregate.index_record_position(), 1_470);
    assert_eq!(aggregate.blocks().len(), 1);
    assert_eq!(
        aggregate
            .block_for_row(0)
            .unwrap()
            .dbcell()
            .record_position(),
        1_696
    );
    assert_eq!(aggregate.first_cell_position(0), Some(1_682));
    assert_eq!(
        &aggregate.to_index_record_bytes().unwrap()[4..],
        &SIMPLE_INDEX
    );
    assert_eq!(
        &aggregate.blocks()[0].to_record_bytes().unwrap()[4..],
        &SIMPLE_DBCELL
    );
}

#[test]
fn rejects_malformed_sizes_bounds_offsets_cardinality_and_targets() {
    assert!(WorksheetIndexRecord::parse_payload(&SIMPLE_INDEX[..19], 2_500).is_err());
    let mut bad = SIMPLE_INDEX;
    bad[0] = 1;
    assert!(WorksheetIndexRecord::parse_payload(&bad, 2_500).is_err());
    let mut bad = SIMPLE_INDEX;
    bad[8..12].copy_from_slice(&0u32.to_le_bytes());
    assert!(WorksheetIndexRecord::parse_payload(&bad, 2_500).is_err());
    let mut bad = SIMPLE_INDEX;
    bad[12..16].copy_from_slice(&2_500u32.to_le_bytes());
    assert!(WorksheetIndexRecord::parse_payload(&bad, 2_500).is_err());
    let mut too_many = vec![0u8; INDEX_FIXED_LEN + (MAX_ROW_BLOCKS + 1) * 4];
    too_many[8..12].copy_from_slice(&1u32.to_le_bytes());
    assert!(WorksheetIndexRecord::parse_payload(&too_many, 20_000).is_err());

    assert!(DbCellRecord::parse_payload(100, 200, &[0, 0, 0]).is_err());
    assert!(DbCellRecord::parse_payload(100, 200, &[0; 69]).is_err());
    assert!(DbCellRecord::parse_payload(20, 200, &[21, 0, 0, 0]).is_err());
    let dbcell = DbCellRecord::parse_payload(100, 200, &[10, 0, 0, 0, 50, 0]).unwrap();
    assert!(dbcell.resolve_cell_positions(60).is_err());

    let mut collector = simple_collector();
    collector.index.as_mut().unwrap().1.dbcell_positions[0] = 1_697;
    assert!(collector.finish().is_err());
    let mut collector = simple_collector();
    collector
        .index
        .as_mut()
        .unwrap()
        .1
        .default_column_width_position = 1_639;
    assert!(collector.finish().is_err());
    let mut collector = RowBlockIndexCollector::new(2_500, 1_450);
    collector.feed_record(1_450, BOF_RECORD_TYPE, &[0; 16]);
    collector.feed_record(1_470, INDEX_RECORD_TYPE, &SIMPLE_INDEX);
    collector.feed_record(1_638, DEF_COL_WIDTH_RECORD_TYPE, &[8, 0]);
    collector.feed_record(1_662, ROW_RECORD_TYPE, &SIMPLE_ROW);
    collector.feed_record(1_683, 0x00FD, &SIMPLE_LABEL_SST);
    collector.feed_record(1_696, DBCELL_RECORD_TYPE, &SIMPLE_DBCELL);
    collector.feed_record(1_750, EOF_RECORD_TYPE, &[]);
    assert!(collector.finish().is_err());
}
