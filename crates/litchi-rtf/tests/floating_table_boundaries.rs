#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{RtfDocument, RtfWriter};

#[test]
fn splits_real_libreoffice_floating_tables_at_non_table_paragraphs() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
        "../../test-data/libreoffice-core/sw/qa/writerfilter/rtftok/data/floattable-tbl-overlap.rtf",
    );
    let document = RtfDocument::parse(&std::fs::read_to_string(path).unwrap()).unwrap();
    assert_eq!(document.tables().len(), 2);
    assert_eq!(document.tables()[0].row_count(), 2);
    assert_eq!(document.tables()[1].row_count(), 2);
    for table in document.tables() {
        assert!(
            table
                .rows()
                .windows(2)
                .all(|rows| rows[0].positioning() == rows[1].positioning())
        );
    }
    assert!(!document.tables()[0].rows()[0].positioning().no_overlap);
    assert!(document.tables()[1].rows()[0].positioning().no_overlap);
}

#[test]
fn formatting_only_gap_is_an_ambiguous_continuation() {
    let source = r"{\rtf1\trowd\tphpg\tposx10\cellx1\intbl A\cell\row \pard {\b } \trowd\tphpg\tposx10\cellx1\intbl B\cell\row}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.tables().len(), 1);
    assert_eq!(document.tables()[0].row_count(), 2);
}

#[test]
fn rejects_inconsistent_positioning_without_a_provable_boundary() {
    for source in [
        r"{\rtf1\trowd\tphpg\tposx10\cellx1\intbl A\cell\row\trowd\tphpg\tposx20\cellx1\intbl B\cell\row}",
        r"{\rtf1\trowd\tphpg\tposx10\cellx1\intbl A\cell\row\pard {\*\unknown\tposx20 ignored}\trowd\tphpg\tposx20\cellx1\intbl B\cell\row}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn writer_preserves_logical_boundaries_deterministically() {
    let source = r"{\rtf1\trowd\tphpg\tposx10\cellx1\intbl A\cell\row\pard Between\par\trowd\tpvpara\tposy20\tabsnoovrlp\cellx1\intbl B\cell\row}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.tables().len(), 2);
    let mut first = Vec::new();
    RtfWriter::new(&mut first)
        .write_document(&document)
        .unwrap();
    let first = String::from_utf8(first).unwrap();
    assert!(first.contains("\\row\n\\pard\\par\n\\trowd"));
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(reparsed.tables().len(), 2);
    assert_eq!(
        reparsed.tables()[0].rows()[0].positioning(),
        document.tables()[0].rows()[0].positioning()
    );
    assert_eq!(
        reparsed.tables()[1].rows()[0].positioning(),
        document.tables()[1].rows()[0].positioning()
    );
    let mut second = Vec::new();
    RtfWriter::new(&mut second)
        .write_document(&reparsed)
        .unwrap();
    assert_eq!(first, String::from_utf8(second).unwrap());
}

#[test]
fn caps_inferred_logical_table_count() {
    let mut source = String::from("{\\rtf1");
    for index in 0..4097 {
        source.push_str("\\trowd\\cellx1\\intbl X\\cell\\row");
        if index < 4096 {
            source.push_str("\\pard S\\par");
        }
    }
    source.push('}');
    assert!(RtfDocument::parse(&source).is_err());
}
