#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::{Package, edit::Error};

fn source() -> Package {
    Package::parse(r"{\rtf1\ansi First\par Second\par Third\par Fourth}").unwrap()
}

fn paragraphs(document: &Package) -> Vec<String> {
    document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect()
}

fn assert_paragraphs(document: &Package, expected: &[&str]) {
    assert_eq!(
        paragraphs(document),
        expected
            .iter()
            .map(|text| (*text).to_owned())
            .collect::<Vec<_>>()
    );
}

#[test]
fn batch_positions_are_relative_to_the_original_snapshot() {
    let source = source();
    let mut edit = source.edit();
    edit.remove_body_paragraphs(&[Position::new(0), Position::new(2)])
        .unwrap();
    let commit = edit.commit().unwrap();

    assert_paragraphs(commit.package(), &["Second", "Fourth"]);
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi Second\par Fourth}"
    );
    assert!(
        source
            .apply(commit.patch())
            .unwrap()
            .into_package()
            .same_snapshot(commit.package())
    );
    assert!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .same_snapshot(&source)
    );
}

#[test]
fn first_middle_last_and_all_removals_preserve_the_unselected_story() {
    for (positions, expected) in [
        (&[Position::new(0)][..], &["Second", "Third", "Fourth"][..]),
        (&[Position::new(1)][..], &["First", "Third", "Fourth"][..]),
        (&[Position::new(3)][..], &["First", "Second", "Third"][..]),
        (
            &[
                Position::new(0),
                Position::new(1),
                Position::new(2),
                Position::new(3),
            ][..],
            &[][..],
        ),
    ] {
        let source = source();
        let mut edit = source.edit();
        edit.remove_body_paragraphs(positions).unwrap();
        let commit = edit.commit().unwrap();
        assert_paragraphs(commit.package(), expected);
        assert_eq!(commit.diagnostics().operation_count(), positions.len());
    }
}

#[test]
fn scalar_and_single_batch_removal_share_source_boundaries() {
    let source = source();

    let mut scalar = source.edit();
    scalar.remove_paragraph(Position::new(3)).unwrap();
    let scalar = scalar.commit().unwrap();

    let mut batch = source.edit();
    batch.remove_body_paragraphs(&[Position::new(3)]).unwrap();
    let batch = batch.commit().unwrap();

    assert_eq!(
        scalar.package().to_bytes().unwrap(),
        batch.package().to_bytes().unwrap()
    );
    assert_paragraphs(scalar.package(), &["First", "Second", "Third"]);
}

#[test]
fn adjacent_empty_paragraph_removal_is_one_atomic_source_splice() {
    let source = Package::parse(r"{\rtf1\ansi First\par\par Third}").unwrap();
    let mut edit = source.edit();
    edit.remove_body_paragraphs(&[Position::new(1)]).unwrap();
    let commit = edit.commit().unwrap();

    assert_paragraphs(commit.package(), &["First", "Third"]);
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi First\par Third}"
    );
}

#[test]
fn trailing_paragraph_boundary_supports_empty_middle_removal() {
    let source = Package::parse(r"{\rtf1\ansi First\par\par Third\par}").unwrap();
    let mut edit = source.edit();
    edit.remove_body_paragraphs(&[Position::new(1)]).unwrap();
    let commit = edit.commit().unwrap();

    assert_paragraphs(commit.package(), &["First", "Third"]);
    assert_eq!(
        commit.package().to_bytes().unwrap(),
        br"{\rtf1\ansi First\par Third\par}"
    );
}

#[test]
fn batch_removal_keeps_an_unselected_empty_paragraph() {
    let source = Package::parse(r"{\rtf1\ansi First\par\par Third}").unwrap();

    let mut ends = source.edit();
    ends.remove_body_paragraphs(&[Position::new(2)]).unwrap();
    let ends = ends.commit().unwrap();
    assert_paragraphs(ends.package(), &["First", ""]);

    let mut separated = source.edit();
    separated
        .remove_body_paragraphs(&[Position::new(0), Position::new(2)])
        .unwrap();
    let separated = separated.commit().unwrap();
    assert_paragraphs(separated.package(), &[""]);
}

#[test]
fn removal_batch_validates_empty_order_and_bounds_without_staging() {
    let source = source();

    let mut empty = source.edit();
    assert!(matches!(
        empty.remove_body_paragraphs(&[]),
        Err(Error::EmptyParagraphRemovalBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    let mut unordered = source.edit();
    assert!(matches!(
        unordered.remove_body_paragraphs(&[Position::new(2), Position::new(1)]),
        Err(Error::ParagraphRemovalBatchOutOfOrder { .. })
    ));
    assert_eq!(unordered.operation_count(), 0);

    let mut late = source.edit();
    assert!(matches!(
        late.remove_body_paragraphs(&[Position::new(0), Position::new(99)]),
        Err(Error::ParagraphOutOfRange {
            position: 99,
            count: 4
        })
    ));
    assert_eq!(late.operation_count(), 0);
}

#[test]
fn protected_batch_refuses_publication_and_leaves_source_unchanged() {
    let source = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1 First\par Second}").unwrap();
    let exact = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.remove_body_paragraphs(&[Position::new(0), Position::new(1)])
        .unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));
    assert_eq!(source.to_bytes().unwrap(), exact);
}

#[test]
fn batch_patch_inverse_is_exact_and_durable_projection_is_source_bound() {
    let source = source();
    let mut edit = source.edit();
    edit.remove_body_paragraphs(&[Position::new(1), Position::new(3)])
        .unwrap();
    let commit = edit.commit().unwrap();
    let limits = litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        8,
        8,
        256 * 1024,
        512 * 1024,
    );
    assert!(matches!(
        commit.patch().to_durable(limits),
        Err(litchi_core::patch::PatchError::InvalidText { .. })
    ));
}
