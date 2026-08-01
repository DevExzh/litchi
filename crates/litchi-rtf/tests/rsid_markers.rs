//! Round-trip tests for in-content RSID markers:
//! `\insrsid`, `\delrsid`, `\charrsid`, `\pararsid`, `\sectrsid`, `\tblrsid`.

use litchi_rtf::{RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn content_rsid_markers_reach_the_model() {
    let source = r#"{\rtf1\ansi\sectrsid46 A\insrsid42 B\delrsid43 C\charrsid44 D\pararsid45\par\trowd\tblrsid47\cellx2000\pard\intbl X\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();

    let blocks = document.blocks();
    let inserted = blocks
        .iter()
        .find(|block| block.text.contains('B'))
        .expect("inserted run");
    assert_eq!(inserted.formatting.insert_rsid, Some(42));
    assert_eq!(inserted.formatting.delete_rsid, Some(43));
    assert_eq!(inserted.formatting.char_style_rsid, Some(44));
    assert_eq!(inserted.paragraph.paragraph_rsid, Some(45));

    assert_eq!(document.sections()[0].properties.section_rsid, Some(46));
    assert_eq!(document.tables()[0].rows()[0].table_rsid(), Some(47));
}

#[test]
fn content_rsid_markers_round_trip_through_the_writer() {
    let source = r#"{\rtf1\ansi\sectrsid46 A\insrsid42 B\delrsid43 C\charrsid44 D\pararsid45\par\trowd\tblrsid47\cellx2000\pard\intbl X\cell\row}"#;
    let document = RtfDocument::parse(source).unwrap();

    let output = write(&document);
    let serialized = String::from_utf8(output).unwrap();
    for marker in [
        r"\insrsid42",
        r"\delrsid43",
        r"\charrsid44",
        r"\pararsid45",
        r"\sectrsid46",
        r"\tblrsid47",
    ] {
        assert!(
            serialized.contains(marker),
            "missing {marker} in {serialized}"
        );
    }

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    let inserted = reparsed
        .blocks()
        .iter()
        .find(|block| block.text.contains('B'))
        .expect("inserted run");
    assert_eq!(inserted.formatting.insert_rsid, Some(42));
    assert_eq!(inserted.formatting.delete_rsid, Some(43));
    assert_eq!(inserted.formatting.char_style_rsid, Some(44));
    assert_eq!(inserted.paragraph.paragraph_rsid, Some(45));
    assert_eq!(reparsed.sections()[0].properties.section_rsid, Some(46));
    assert_eq!(reparsed.tables()[0].rows()[0].table_rsid(), Some(47));
}
