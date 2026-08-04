use litchi_ooxml::docx::Package;
use std::io::Cursor;

#[test]
fn reopened_document_statistics_use_the_concrete_metrics_owner() {
    let mut package = Package::new().unwrap();
    {
        let document = package.document_mut().unwrap();
        document.add_paragraph_with_text("alpha beta");
        document.add_paragraph();
    }

    let mut bytes = Cursor::new(Vec::new());
    package.to_stream(&mut bytes).unwrap();

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
