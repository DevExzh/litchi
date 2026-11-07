use litchi_docx::Package;
use std::io::Cursor;

#[test]
fn reopened_document_statistics_use_the_concrete_metrics_owner() {
    let mut builder = litchi_docx::Builder::new().unwrap();
    {
        let document = builder.document().unwrap();
        document.add_paragraph_with_text("alpha beta");
        document.add_paragraph();
    }
    let package = builder.build().unwrap();

    let mut bytes = Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();

    let reopened = Package::from_reader(Cursor::new(bytes.into_inner())).unwrap();
    let statistics = reopened.document().unwrap().statistics().unwrap();

    assert_eq!(statistics.word_count(), 2);
    assert_eq!(statistics.character_count(), 10);
    assert_eq!(statistics.character_count_no_spaces(), 9);
    assert_eq!(statistics.paragraph_count(), 2);
    assert_eq!(statistics.line_count(), 1);
    assert_eq!(statistics.page_count(), 1);
    assert_eq!(statistics.table_count(), 0);
    assert_eq!(statistics.image_count(), 0);
    assert_eq!(statistics.drawing_count(), 0);
}

#[test]
fn paragraph_and_table_cell_boundaries_separate_words_but_run_boundaries_do_not() {
    let mut builder = litchi_docx::Builder::new().unwrap();
    let doc = builder.document().unwrap();
    let p = doc.add_paragraph();
    p.add_run_with_text("con");
    p.add_run_with_text("tiguous");
    doc.add_paragraph_with_text("next");
    let table = doc.add_table(1, 2);
    table.cell(0, 0).unwrap().set_text("table");
    table.cell(0, 1).unwrap().set_text("cell");
    let p = builder.build().unwrap();
    let d = p.document().unwrap();
    let s = d.statistics().unwrap();
    assert_eq!(d.text().unwrap(), "contiguousnexttablecell");
    assert_eq!(s.word_count(), 4);
    assert_eq!(s.character_count(), 23);
    assert_eq!(s.paragraph_count(), 4);
    assert_eq!(s.table_count(), 1);
}
