#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

//! Round-trip tests for structural revision metadata: `\prauthN`/`\prdateN`,
//! `\srauthN`/`\srdateN`, `\trauthN`/`\trdateN`, and the `\clins`, `\cldel`,
//! `\clmrgd` cell revision markers with their author/DTTM companions.

use litchi_rtf::{CellRevisionKind, RevisionMetadata, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

fn metadata(author: i32, date: i32) -> RevisionMetadata {
    RevisionMetadata {
        author: Some(author),
        date: Some(date),
    }
}

#[test]
fn paragraph_revision_metadata_round_trips() {
    let source = r"{\rtf1\ansi\prauth2\prdate1199059860 Revised paragraph\par}";
    let document = RtfDocument::parse(source).unwrap();
    let block = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Revised"))
        .expect("paragraph block");
    assert_eq!(block.paragraph.revision, metadata(2, 1_199_059_860));

    let first = write(&document);
    assert!(first.contains("\\prauth2"), "missing prauth in {first}");
    assert!(
        first.contains("\\prdate1199059860"),
        "missing prdate in {first}"
    );
    let reparsed = RtfDocument::parse(&first).unwrap();
    let block = reparsed
        .blocks()
        .iter()
        .find(|block| block.text.contains("Revised"))
        .expect("paragraph block");
    assert_eq!(block.paragraph.revision, metadata(2, 1_199_059_860));
    assert_eq!(first, write(&reparsed));
}

#[test]
fn paragraph_revision_metadata_resets_with_pard() {
    let source = r"{\rtf1\ansi\prauth1\prdate42 One\par\pard Two\par}";
    let document = RtfDocument::parse(source).unwrap();
    let two = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Two"))
        .expect("second paragraph");
    assert_eq!(two.paragraph.revision, RevisionMetadata::default());
}

#[test]
fn negative_paragraph_revision_author_is_rejected() {
    assert!(RtfDocument::parse(r"{\rtf1\ansi\prauth-1 X\par}").is_err());
}

#[test]
fn section_revision_metadata_round_trips() {
    let source = r"{\rtf1\ansi\sectd\srauth3\srdate-1501115711 Body\par}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(
        document.sections()[0].properties.revision,
        metadata(3, -1_501_115_711)
    );

    let first = write(&document);
    assert!(first.contains("\\srauth3"), "missing srauth in {first}");
    assert!(
        first.contains("\\srdate-1501115711"),
        "missing srdate in {first}"
    );
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(
        reparsed.sections()[0].properties.revision,
        metadata(3, -1_501_115_711)
    );
    assert_eq!(first, write(&reparsed));
}

#[test]
fn table_row_revision_metadata_round_trips() {
    let source = r"{\rtf1\trowd\trauth4\trdate777 \cellx1000\intbl A\cell\row\trowd\cellx1000\intbl B\cell\row}";
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(rows[0].revision(), metadata(4, 777));
    // \trowd resets row revision metadata for the next row.
    assert_eq!(rows[1].revision(), RevisionMetadata::default());

    let first = write(&document);
    assert!(first.contains("\\trauth4"), "missing trauth in {first}");
    assert!(first.contains("\\trdate777"), "missing trdate in {first}");
    let reparsed = RtfDocument::parse(&first).unwrap();
    assert_eq!(reparsed.tables()[0].rows()[0].revision(), metadata(4, 777));
    assert_eq!(first, write(&reparsed));
}

#[test]
fn cell_revision_marks_and_metadata_round_trip() {
    let source = concat!(
        r"{\rtf1\trowd",
        r"\clins\clinsauth1\clinsdttm100\cellx900",
        r"\cldel\cldelauth2\cldeldttm200\cellx1800",
        r"\clmrgd\clmrgdauth3\clmrgddttm300\cellx2700",
        r"\intbl A\cell B\cell C\cell\row}"
    );
    let document = RtfDocument::parse(source).unwrap();
    let cells = document.tables()[0].rows()[0].cells();
    for (cell, kind, author, date) in [
        (&cells[0], CellRevisionKind::Inserted, 1, 100),
        (&cells[1], CellRevisionKind::Deleted, 2, 200),
        (&cells[2], CellRevisionKind::MergeDeleted, 3, 300),
    ] {
        let revision = cell.revision().expect("cell revision");
        assert_eq!(revision.kind, kind);
        assert_eq!(revision.metadata, metadata(author, date));
    }

    let first = write(&document);
    for expected in [
        "\\clins\\clinsauth1\\clinsdttm100",
        "\\cldel\\cldelauth2\\cldeldttm200",
        "\\clmrgd\\clmrgdauth3\\clmrgddttm300",
    ] {
        assert!(first.contains(expected), "missing {expected} in {first}");
    }
    let reparsed = RtfDocument::parse(&first).unwrap();
    let cells = reparsed.tables()[0].rows()[0].cells();
    assert_eq!(
        cells[0].revision().expect("cell revision").kind,
        CellRevisionKind::Inserted
    );
    assert_eq!(first, write(&reparsed));
}

#[test]
fn cell_revision_metadata_requires_matching_mark() {
    // Author/DTTM companions require a preceding \clins/\cldel/\clmrgd.
    assert!(RtfDocument::parse(r"{\rtf1\trowd\clinsauth1\cellx900\intbl A\cell\row}").is_err());
    assert!(
        RtfDocument::parse(r"{\rtf1\trowd\cldel\clinsdttm1\cellx900\intbl A\cell\row}").is_err()
    );
    // Conflicting revision markers on one cell are rejected.
    assert!(RtfDocument::parse(r"{\rtf1\trowd\clins\cldel\cellx900\intbl A\cell\row}").is_err());
    // Negative author indices are rejected.
    assert!(
        RtfDocument::parse(r"{\rtf1\trowd\clins\clinsauth-1\cellx900\intbl A\cell\row}").is_err()
    );
}

#[test]
fn cell_revision_resets_per_cell_and_row() {
    let source = concat!(
        r"{\rtf1\trowd\clins\cellx900\cellx1800\intbl A\cell B\cell\row",
        r"\trowd\cellx900\intbl C\cell\row}"
    );
    let document = RtfDocument::parse(source).unwrap();
    let rows = document.tables()[0].rows();
    assert_eq!(
        rows[0].cells()[0].revision().map(|revision| revision.kind),
        Some(CellRevisionKind::Inserted)
    );
    // The second cellx in the same row starts a fresh cell definition.
    assert_eq!(rows[0].cells()[1].revision(), None);
    // \trowd clears pending cell revisions for the next row.
    assert_eq!(rows[1].cells()[0].revision(), None);

    let first = write(&document);
    assert_eq!(first, write(&RtfDocument::parse(&first).unwrap()));
}

#[test]
fn nested_table_cell_and_row_revision_round_trip() {
    let source = concat!(
        r"{\rtf1\trowd\cellx1000\intbl\itap1 outer ",
        r"\intbl\itap2 Inner\nestcell",
        r"{\*\nesttableprops\trowd\trauth1\trdate9\cldel\cldelauth2\cldeldttm8\cellx500\nestrow}",
        r"\intbl\itap1 tail\cell\row}"
    );
    let document = RtfDocument::parse(source).unwrap();
    let outer = &document.tables()[0];
    let nested = &outer.rows()[0].cells()[0].nested_tables()[0].table;
    let row = &nested.rows()[0];
    assert_eq!(row.revision(), metadata(1, 9));
    let revision = row.cells()[0].revision().expect("nested cell revision");
    assert_eq!(revision.kind, CellRevisionKind::Deleted);
    assert_eq!(revision.metadata, metadata(2, 8));

    let first = write(&document);
    assert_eq!(first, write(&RtfDocument::parse(&first).unwrap()));
    assert!(first.contains("\\trauth1"), "missing trauth in {first}");
    assert!(first.contains("\\cldel"), "missing cldel in {first}");
}
