//! Regression coverage for the allocation-free native-DOC paragraph count.

use litchi_doc::Package;
use litchi_doc::writer::Writer;
use std::io::Cursor;

#[test]
fn paragraph_count_matches_independent_materialized_oracle() {
    let mut writer = Writer::new();
    for text in ["first", "second 😀", "third"] {
        writer.add_paragraph(text).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();

    assert_eq!(document.paragraph_count().unwrap(), paragraphs.len());
    assert_eq!(paragraphs.len(), 3);
    assert_eq!(document.text().unwrap(), "first\rsecond 😀\rthird\r");
}
