use std::io::Cursor;

use litchi_xls::Workbook;
use litchi_xls::writer::Writer;

#[test]
fn generated_worksheet_indexes_regenerate_and_parse() {
    let mut writer = Writer::new();
    writer.add_worksheet("Empty").unwrap();

    let ordinary = writer.add_worksheet("Ordinary").unwrap();
    writer.write_number(ordinary, 0, 0, 1.0).unwrap();
    writer.write_string(ordinary, 4, 2, "indexed").unwrap();

    let boundary = writer.add_worksheet("Boundary").unwrap();
    for row in 0..33 {
        writer
            .write_number(boundary, row, 0, f64::from(row))
            .unwrap();
    }

    let sparse = writer.add_worksheet("Sparse").unwrap();
    writer.write_number(sparse, 1, 0, 1.0).unwrap();
    writer.write_number(sparse, 76, 0, 76.0).unwrap();

    let formatting = writer.add_worksheet("Formatting").unwrap();
    writer.set_row_height(formatting, 0, 18.0).unwrap();
    writer.set_row_height(formatting, 40, 18.0).unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();

    let empty = workbook
        .xls_worksheet(0)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(empty.blocks().len(), 0);

    let ordinary = workbook
        .xls_worksheet(1)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(ordinary.index_record().first_data_row(), 0);
    assert_eq!(ordinary.index_record().last_data_row_exclusive(), 5);
    assert_eq!(ordinary.blocks().len(), 1);

    let boundary = workbook
        .xls_worksheet(2)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(boundary.blocks().len(), 2);

    let sparse = workbook
        .xls_worksheet(3)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(sparse.index_record().first_data_row(), 1);
    assert_eq!(sparse.index_record().last_data_row_exclusive(), 77);
    assert_eq!(sparse.blocks().len(), 3);
    assert!(sparse.blocks()[1].indexed_rows().is_empty());

    let formatting = workbook
        .xls_worksheet(4)
        .unwrap()
        .row_block_index()
        .unwrap()
        .unwrap();
    assert_eq!(formatting.blocks().len(), 2);
    assert!(formatting.blocks()[0].indexed_rows().is_empty());
    assert!(formatting.blocks()[1].indexed_rows().is_empty());
}
