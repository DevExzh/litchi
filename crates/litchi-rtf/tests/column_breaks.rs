#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::{BodyStoryEvent, CellStoryEvent, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn body_column_is_zero_width_ordered_and_round_trips_canonically() {
    let document = Tree::parse(
        r"{\rtf1\ansi A\column B}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "AB");
    assert!(matches!(
        document.body_story_events(),
        [BodyStoryEvent::ColumnBreak(column)] if column.position == 1
    ));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\column "));
    assert!(!serialized.contains("A\\par B"));
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.body_story_events(), document.body_story_events());
}

#[test]
fn table_and_nested_table_column_events_round_trip_in_place() {
    let source = r"{\rtf1\trowd\cellx5000\intbl\itap1 Before\column After \intbl\itap2 Inner\column Tail\nestcell{\*\nesttableprops\itap2\trowd\cellx1000\nestrow}{\nonesttables\par}\intbl\itap1 End\cell\row}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let outer = &document.tables()[0].rows()[0].cells()[0];
    assert!(
        outer.story_events().iter().any(
            |event| matches!(event, CellStoryEvent::ColumnBreak(column) if column.position == 6),
        )
    );
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert!(
        inner.story_events().iter().any(
            |event| matches!(event, CellStoryEvent::ColumnBreak(column) if column.position == 5),
        )
    );

    let output = write(&document);
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    let outer = &reparsed.tables()[0].rows()[0].cells()[0];
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert!(
        outer
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::ColumnBreak(_)))
    );
    assert!(
        inner
            .story_events()
            .iter()
            .any(|event| matches!(event, CellStoryEvent::ColumnBreak(_)))
    );
}

#[test]
fn document_column_breaks_can_be_pushed_and_cleared() {
    let mut document =
        Tree::parse(r"{\rtf1\ansi AB}", litchi_rtf::read::OpenOptions::default()).unwrap();
    document.push_column_break(1).unwrap();
    assert!(matches!(
        document.column_breaks().next(),
        Some(column) if column.position == 1
    ));
    assert!(document.push_column_break(3).is_err());
    document.clear_column_breaks();
    assert!(document.column_breaks().next().is_none());
}

#[test]
fn rejects_parameters_in_the_body() {
    assert!(repair_support::rejected_or_parameter_repaired(
        &Tree::parse(
            r"{\rtf1 A\column1 B}",
            litchi_rtf::read::OpenOptions::default()
        )
    ));
}

#[test]
fn parses_the_repository_column_break_fixture() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/rtf/column-break.rtf");
    let bytes = std::fs::read(&path).unwrap();
    let document = Tree::from_bytes(&bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(matches!(
        document.body_story_events(),
        [BodyStoryEvent::ColumnBreak(column)] if column.position == 0
    ));
}
