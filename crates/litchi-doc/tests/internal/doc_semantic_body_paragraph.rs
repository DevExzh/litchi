//! Regression coverage for borrowed ordinary-body paragraph queries.

use litchi_core::Position;
use litchi_doc::body_text::{Projection, Snapshot};
use litchi_doc::raw::Reader;
use litchi_doc::writer::{Builder, FootnoteEntry};
use std::io::Cursor;

fn write_fixture() -> Vec<u8> {
    let mut writer = Builder::new();
    writer
        .add_paragraph("ordinary")
        .expect("ordinary paragraph");
    writer.add_paragraph("").expect("blank paragraph");
    writer
        .add_paragraph("Unicode 😀 and 𝄞")
        .expect("Unicode paragraph");
    let table = writer.add_table(1, 1).expect("table");
    writer
        .set_table_cell_text(table, 0, 0, "table cell")
        .expect("table cell text");
    writer.set_odd_header("header story");
    writer.add_footnote(FootnoteEntry::new(0, "footnote story", 1));

    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("DOC fixture builds")
        .write_to(&mut output)
        .expect("DOC fixture");
    output.into_inner()
}

#[test]
fn body_queries_match_native_source_projection_and_ignore_other_stories() {
    let bytes = write_fixture();
    let native = Snapshot::parse(&bytes)
        .expect("native body snapshot")
        .paragraphs(Projection::All)
        .expect("native body paragraphs");

    let mut package = Reader::from_reader(Cursor::new(bytes)).expect("DOC package");
    let document = package.document().expect("DOC document");

    assert_eq!(
        document
            .body_paragraph_count()
            .expect("body paragraph count"),
        native.len()
    );
    for (position, paragraph) in native.iter().enumerate() {
        assert_eq!(
            document
                .body_paragraph_text(Position::new(position))
                .expect("body paragraph query"),
            Some(paragraph.text())
        );
    }
    assert_eq!(
        document
            .body_paragraph_text(Position::new(native.len()))
            .expect("out-of-range body paragraph query"),
        None
    );

    let mut iterated = document.body_paragraphs();
    for paragraph in &native {
        assert_eq!(
            iterated
                .next()
                .expect("paragraph")
                .expect("valid paragraph"),
            paragraph.text()
        );
    }
    assert!(iterated.next().is_none());
}
