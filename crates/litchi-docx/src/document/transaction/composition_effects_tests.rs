use litchi_core::Position;

use super::super::{RevisionKind, Snapshot, SubEditJoinFailure};
use super::{CompositionLimits, operation_effects};

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn document(body: &str) -> Vec<u8> {
    format!("<w:document xmlns:w=\"{WORD}\"><w:body>{body}<w:sectPr/></w:body></w:document>")
        .into_bytes()
}

fn limits() -> CompositionLimits {
    CompositionLimits::new(8, 16, 64, 16)
}

#[test]
fn revision_action_conflicts_with_same_paragraph_source_owner() {
    let source = Snapshot::from_xml(document(
        "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>",
    ))
    .unwrap();

    let mut action = source.edit();
    action
        .accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let action = action.prepare(limits(), "accept").unwrap();

    let mut text = source.edit();
    text.replace_revision_text(
        Position::new(0),
        RevisionKind::Insertion,
        Position::new(0),
        "rewritten",
    )
    .unwrap();
    let text = text.prepare(limits(), "text").unwrap();

    let mut composition = source.compose(limits());
    composition.join(action).unwrap();
    assert!(matches!(
        composition.join(text).unwrap_err().failure(),
        SubEditJoinFailure::Overlap(_)
    ));
    assert_eq!(composition.len(), 1);
}

#[test]
fn revision_actions_in_separate_paragraphs_remain_disjoint_and_inverse_symmetric() {
    let source = Snapshot::from_xml(document(
        "<w:p><w:ins w:id=\"1\" w:author=\"A\"><w:r><w:t>first</w:t></w:r></w:ins></w:p><w:p><w:ins w:id=\"2\" w:author=\"B\"><w:r><w:t>second</w:t></w:r></w:ins></w:p>",
    ))
    .unwrap();

    let mut first = source.edit();
    first
        .accept_revision(Position::new(0), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let first = first.prepare(limits(), "first").unwrap();
    let operation = first.operations()[0].clone();
    let inverse = operation.inverse();
    assert_eq!(
        operation_effects(std::slice::from_ref(&operation), source.paragraph_count()),
        operation_effects(std::slice::from_ref(&inverse), source.paragraph_count())
    );

    let mut second = source.edit();
    second
        .accept_revision(Position::new(1), RevisionKind::Insertion, Position::new(0))
        .unwrap();
    let second = second.prepare(limits(), "second").unwrap();

    let mut composition = source.compose(limits());
    composition.join(first).unwrap().join(second).unwrap();
    let committed = composition.commit().unwrap();
    let xml = std::str::from_utf8(committed.snapshot().xml_bytes()).unwrap();
    assert!(xml.contains("<w:t>first</w:t>"));
    assert!(xml.contains("<w:t>second</w:t>"));
}
