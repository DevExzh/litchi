use super::super::support::*;

#[test]
fn test_add_table() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 3).unwrap();
    assert_eq!(idx, 0);
    assert_eq!(writer.tables[0].rows.len(), 2);
    assert_eq!(writer.tables[0].rows[0].cells.len(), 3);
}

#[test]
fn test_set_table_cell() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    writer.set_table_cell_text(idx, 0, 0, "Cell").unwrap();
    assert_eq!(
        writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
        "Cell"
    );
}

#[test]
fn test_set_table_cell_multiple() {
    let mut writer = Writer::new();
    let idx = writer.add_table(2, 2).unwrap();
    writer.set_table_cell_text(idx, 0, 0, "A").unwrap();
    writer.set_table_cell_text(idx, 0, 1, "B").unwrap();
    writer.set_table_cell_text(idx, 1, 0, "C").unwrap();
    writer.set_table_cell_text(idx, 1, 1, "D").unwrap();
    assert_eq!(
        writer.tables[0].rows[0].cells[0].paragraphs[0].runs[0].text,
        "A"
    );
    assert_eq!(
        writer.tables[0].rows[0].cells[1].paragraphs[0].runs[0].text,
        "B"
    );
    assert_eq!(
        writer.tables[0].rows[1].cells[0].paragraphs[0].runs[0].text,
        "C"
    );
    assert_eq!(
        writer.tables[0].rows[1].cells[1].paragraphs[0].runs[0].text,
        "D"
    );
}
