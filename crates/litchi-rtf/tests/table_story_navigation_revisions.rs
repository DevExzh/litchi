#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    CellStoryEvent, IndexEntry, NavigationEntry, Revision, RevisionAuthor, RevisionType,
    RtfDocument, RtfWriter, TableCellCoordinate, TableCellPath,
};
use std::borrow::Cow;

fn round_trip(source: &str) -> RtfDocument<'static> {
    let document = RtfDocument::parse(source).unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    RtfDocument::parse_bytes(&output).unwrap()
}

#[test]
fn outer_cell_retains_navigation_and_both_revision_kinds() {
    let source = concat!(
        r#"{\rtf1\ansi{\*\revtbl{Ada;}}Body"#,
        r#"\trowd\cellx1440{\intbl A{\tc\v\tcl2 Cell}"#,
        r#"{\revised\revauth0\revdttm11 new}"#,
        r#"{\deleted\revauthdel0\revdttmdel12 old}Z\cell}\row}"#,
    );
    let document = round_trip(source);
    assert_eq!(document.text(), "Body");
    assert_eq!(document.navigation_entries().len(), 1);
    assert_eq!(document.revisions().len(), 2);
    let cell = &document.tables()[0].rows()[0].cells()[0];
    assert_eq!(cell.text(), "AnewZ");
    assert!(
        cell.story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::NavigationEntry(_)))
    );
    assert!(
        cell.story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::RevisionStart(_)))
    );
    assert!(
        cell.story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::RevisionDeletion(_)))
    );
    assert!(matches!(
        document.navigation_entries()[0],
        NavigationEntry::TableOfContents(_)
    ));
    assert!(
        document
            .revisions()
            .iter()
            .any(|revision| revision.revision_type == RevisionType::Insertion
                && revision.content == "new")
    );
    assert!(
        document
            .revisions()
            .iter()
            .any(|revision| revision.revision_type == RevisionType::Deletion
                && revision.content == "old")
    );
}

#[test]
fn nested_cell_and_body_metadata_remain_independent() {
    let source = concat!(
        r#"{\rtf1\ansi{\*\revtbl{Ada;}}{\tc Body}"#,
        r#"{\revised\revauth0\revdttm1 body}"#,
        r#"\trowd\cellx1440{\intbl O"#,
        r#"{\intbl\itap2 N{\xe\v nested}{\revised\revauth0\revdttm2 ins}\nestcell}"#,
        r#"{\*\nesttableprops\itap2\trowd\cellx720\nestrow}"#,
        r#"X\cell}\row}"#,
    );
    let document = round_trip(source);
    assert_eq!(document.navigation_entries().len(), 2);
    assert_eq!(document.revisions().len(), 2);
    let outer = &document.tables()[0].rows()[0].cells()[0];
    let nested = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert_eq!(outer.text(), "OX");
    assert_eq!(nested.text(), "Nins");
    assert!(nested.navigation_entry_references().next().is_some());
    assert_eq!(nested.revision_events().count(), 2);
}

#[test]
fn rejects_missing_metadata_conflicts_and_active_destinations() {
    for source in [
        r"{\rtf1\trowd\cellx1000{\intbl{\revised x}\cell}\row}",
        r"{\rtf1{\*\revtbl{A;}}\trowd\cellx1000{\intbl{\revised\deleted\revauth0 x}\cell}\row}",
        r"{\rtf1\trowd\cellx1000{\intbl{\xe x{\field danger}}\cell}\row}",
        r"{\rtf1\trowd\cellx1000{\intbl{\tc\tcl10 bad}\cell}\row}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn writer_rejects_unowned_multiple_wrong_kind_and_out_of_range_mutations() {
    let mut unowned = RtfDocument::parse(r"{\rtf1\trowd\cellx1000{\intbl A\cell}\row}").unwrap();
    unowned
        .push_cell_navigation_entry_metadata(NavigationEntry::Index(
            IndexEntry::new(0, Cow::Borrowed("x")).unwrap(),
        ))
        .unwrap();
    assert!(RtfWriter::new(Vec::new()).write_document(&unowned).is_err());

    let mut duplicate = RtfDocument::parse(
        r"{\rtf1\trowd\cellx1000\cellx2000{\intbl A{\xe\v x}\cell}{\intbl B\cell}\row}",
    )
    .unwrap();
    duplicate.tables_mut()[0].rows_mut()[0].cells_mut()[1]
        .push_navigation_entry_reference(0, 0)
        .unwrap();
    assert!(
        RtfWriter::new(Vec::new())
            .write_document(&duplicate)
            .is_err()
    );

    let mut wrong = RtfDocument::parse(
        r"{\rtf1{\*\revtbl{A;}}\trowd\cellx1000{\intbl{\revised\revauth0 x}\cell}\row}",
    )
    .unwrap();
    let cell = &mut wrong.tables_mut()[0].rows_mut()[0].cells_mut()[0];
    cell.clear_revision_references();
    cell.push_deletion_revision_reference(0, 0).unwrap();
    assert!(RtfWriter::new(Vec::new()).write_document(&wrong).is_err());

    let mut out_of_range =
        RtfDocument::parse(r"{\rtf1\trowd\cellx1000{\intbl A\cell}\row}").unwrap();
    out_of_range.tables_mut()[0].rows_mut()[0].cells_mut()[0]
        .push_navigation_entry_reference(99, 0)
        .unwrap();
    assert!(
        RtfWriter::new(Vec::new())
            .write_document(&out_of_range)
            .is_err()
    );
}

#[test]
fn atomic_outer_and_nested_mutation_round_trips() {
    let mut document = RtfDocument::parse(concat!(
        r#"{\rtf1\trowd\cellx1000{\intbl O"#,
        r#"{\intbl\itap2 N\nestcell}"#,
        r#"{\*\nesttableprops\itap2\trowd\cellx500\nestrow}"#,
        r#"X\cell}\row}"#,
    ))
    .unwrap();
    document
        .push_revision_author(RevisionAuthor::new(Cow::Borrowed("Ada")).unwrap())
        .unwrap();
    let outer = TableCellPath::outer(0, 0, 0);
    document
        .push_navigation_entry_for_cell(
            &outer,
            NavigationEntry::Index(IndexEntry::new(1, Cow::Borrowed("outer")).unwrap()),
        )
        .unwrap();
    let nested = outer.clone().with_nested(TableCellCoordinate {
        table_index: 0,
        row_index: 0,
        cell_index: 0,
    });
    document
        .push_revision_for_cell(
            &nested,
            Revision {
                revision_type: RevisionType::Insertion,
                author: Cow::Borrowed("Ada"),
                date: Some(Cow::Borrowed("7")),
                id: 0,
                content: Cow::Borrowed("N"),
                position: 0,
                range_end: 1,
            },
        )
        .unwrap();
    let reparsed = round_trip(
        &String::from_utf8({
            let mut output = Vec::new();
            RtfWriter::new(&mut output)
                .write_document(&document)
                .unwrap();
            output
        })
        .unwrap(),
    );
    assert_eq!(reparsed.navigation_entries().len(), 1);
    assert_eq!(reparsed.revisions().len(), 1);
    let nested_cell = &reparsed.tables()[0].rows()[0].cells()[0].nested_tables()[0]
        .table
        .rows()[0]
        .cells()[0];
    assert_eq!(nested_cell.revision_events().count(), 2);
}
