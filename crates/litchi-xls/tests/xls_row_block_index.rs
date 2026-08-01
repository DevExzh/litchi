use std::fs::File;

use litchi_xls::XlsWorkbook;

#[test]
fn parses_poi_simple_and_multirow_indexes() {
    let simple_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/Simple.xls"
    );
    let workbook = XlsWorkbook::new(File::open(simple_path).unwrap()).unwrap();
    let index = workbook
        .xls_worksheet(0)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(
        (
            index.index_record().first_data_row(),
            index.index_record().last_data_row_exclusive()
        ),
        (0, 1)
    );
    assert_eq!(index.index_record().default_column_width_position(), 1_638);
    assert_eq!(index.blocks().len(), 1);
    assert_eq!(index.blocks()[0].dbcell().record_position(), 1_696);
    assert_eq!(index.first_cell_position(0), Some(1_682));

    let multi_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/spreadsheet/48968.xls"
    );
    let workbook = XlsWorkbook::new(File::open(multi_path).unwrap()).unwrap();
    let index = workbook
        .xls_worksheet(0)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(
        (
            index.index_record().first_data_row(),
            index.index_record().last_data_row_exclusive()
        ),
        (0, 29)
    );
    assert_eq!(index.blocks().len(), 1);
    assert_eq!(index.blocks()[0].indexed_rows().len(), 24);
    assert_eq!(index.blocks()[0].dbcell().record_position(), 17_918);
    assert_eq!(index.first_cell_position(0), Some(16_491));
    assert_eq!(index.first_cell_position(8), None);
    assert_eq!(index.first_cell_position(28), Some(17_856));
}

#[test]
fn parses_real_sparse_and_rowless_dbcell_blocks() {
    let sparse_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ole/xls/ConditionalFormattingSamples.xls"
    );
    let workbook = XlsWorkbook::new(File::open(sparse_path).unwrap()).unwrap();
    let index = workbook
        .xls_worksheet(13)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(
        (
            index.index_record().first_data_row(),
            index.index_record().last_data_row_exclusive()
        ),
        (1, 92)
    );
    assert_eq!(index.blocks().len(), 3);
    assert!(index.blocks()[1].indexed_rows().is_empty());
    assert!(index.blocks()[1].dbcell().first_row_position().is_none());
    assert!(index.blocks()[1].dbcell().cell_offsets().is_empty());
}
