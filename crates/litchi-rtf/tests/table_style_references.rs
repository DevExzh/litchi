#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{Row, RtfDocument, RtfWriter, StyleType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn retains_group_scope_row_inheritance_last_wins_and_trowd_reset() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\stylesheet{\*\ts5 Declared Table;}}"#,
        r#"\trowd\ts4\ts5{\ts6}\cellx1000\intbl A\cell\row"#,
        r#"\intbl B\cell\row"#,
        r#"\trowd\cellx1000\intbl C\cell\row}"#,
    ))
    .unwrap();
    assert!(
        document
            .stylesheet()
            .get_typed(StyleType::Table, 5)
            .is_some()
    );
    let rows = document.tables()[0].rows();
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[0].table_style(), Some(5));
    assert_eq!(rows[1].table_style(), Some(5));
    assert_eq!(rows[2].table_style(), None);
}

#[test]
fn preserves_zero_maximum_omission_and_public_mutation() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\trowd\ts0\cellx1000\intbl A\cell\row"#,
        r#"\trowd\ts65535\cellx1000\intbl B\cell\row"#,
        r#"\trowd\cellx1000\intbl C\cell\row}"#,
    ))
    .unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].table_style(), Some(0));
    assert_eq!(rows[1].table_style(), Some(65_535));
    assert_eq!(rows[2].table_style(), None);

    let mut row = Row::new();
    row.set_table_style(Some(0));
    assert_eq!(row.table_style(), Some(0));
    row.set_table_style(None);
    assert_eq!(row.table_style(), None);
}

#[test]
fn rejects_malformed_body_and_stylesheet_handles_and_duplicate_selectors() {
    for source in [
        r"{\rtf1\trowd\ts\cellx1000\intbl X\cell\row}",
        r"{\rtf1\trowd\ts-1\cellx1000\intbl X\cell\row}",
        r"{\rtf1\trowd\ts65536\cellx1000\intbl X\cell\row}",
        r"{\rtf1{\stylesheet{\*\ts Missing;}}}",
        r"{\rtf1{\stylesheet{\*\ts-1 Negative;}}}",
        r"{\rtf1{\stylesheet{\*\ts65536 Overflow;}}}",
        r"{\rtf1{\stylesheet{\ts1 Unstarred;}}}",
        r"{\rtf1{\stylesheet{\*\b\ts1 Late;}}}",
        r"{\rtf1{\stylesheet{\*\ts1\ts2 Duplicate;}}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let inert =
        RtfDocument::parse(r"{\rtf1{\field{\*\fldinst TEST \ts65536}{\fldrslt Result}}Body}")
            .unwrap();
    assert!(
        inert
            .tables()
            .iter()
            .flat_map(litchi_rtf::raw::Table::rows)
            .all(|row| row.table_style().is_none())
    );
}

#[test]
fn canonical_writer_round_trips_outer_and_nested_row_references_stably() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\stylesheet{\*\ts3 Declared Table;}}"#,
        r#"\trowd\ts3\cellx5000\intbl\itap2 Inner\nestcell"#,
        r#"{\*\nesttableprops\itap2\trowd\ts4\cellx1000\nestrow}"#,
        r#"\intbl\itap1\cell\row}"#,
    ))
    .unwrap();
    let outer = &document.tables()[0].rows()[0];
    assert_eq!(outer.table_style(), Some(3));
    let nested = &outer.cells()[0].nested_tables()[0].table.rows()[0];
    assert_eq!(nested.table_style(), Some(4));

    let first = write(&document);
    let serialized = String::from_utf8(first.clone()).unwrap();
    assert!(serialized.contains(r"\trowd\ts3"));
    assert!(serialized.contains(r"\trowd\ts4"));
    let reparsed = RtfDocument::parse_bytes(&first).unwrap();
    assert_eq!(reparsed.stylesheet(), document.stylesheet());
    assert_eq!(reparsed.tables()[0].rows()[0].table_style(), Some(3));
    assert_eq!(
        reparsed.tables()[0].rows()[0].cells()[0].nested_tables()[0]
            .table
            .rows()[0]
            .table_style(),
        Some(4)
    );
    assert_eq!(write(&reparsed), first);
}

#[test]
fn parses_libreoffice_table_style_declaration_and_body_references() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf148544.rtf"
    );
    let document = RtfDocument::parse_bytes(source).unwrap();
    assert!(
        document
            .stylesheet()
            .get_typed(StyleType::Table, 11)
            .is_some()
    );
    assert!(
        document
            .tables()
            .iter()
            .flat_map(litchi_rtf::raw::Table::rows)
            .any(|row| row.table_style() == Some(11))
    );
}
