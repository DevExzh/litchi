#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::edit::{Error, ParagraphAlignmentUpdate, TextSpan};
use litchi_rtf::{Alignment, EditOptions, Package, SemanticLimitKind, SemanticLimits};

fn updates(values: &[(usize, Alignment)]) -> Vec<ParagraphAlignmentUpdate> {
    values
        .iter()
        .map(|(position, alignment)| {
            ParagraphAlignmentUpdate::new(Position::new(*position), *alignment)
        })
        .collect()
}

fn alignments(document: &Package) -> Vec<Alignment> {
    document
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.format().alignment())
        .collect()
}

#[test]
fn native_parser_keeps_alignment_on_the_paragraph_before_each_break() {
    let source =
        Package::parse(r"{\rtf1\ansi\qr First\par\qc Second\par\qj Third\par\ql Fourth}").unwrap();

    assert_eq!(
        alignments(&source),
        [
            Alignment::Right,
            Alignment::Center,
            Alignment::Justify,
            Alignment::Left,
        ]
    );
}

#[test]
fn batch_updates_first_middle_and_last_and_reopens() {
    let source = Package::parse(r"{\rtf1\ansi\qc\b First\b0\par\qr Second\par\qj Third}").unwrap();
    let requested = updates(&[
        (0, Alignment::Right),
        (1, Alignment::Center),
        (2, Alignment::Left),
    ]);

    let mut edit = source.edit();
    edit.set_body_paragraph_alignments(&requested).unwrap();
    assert_eq!(edit.operation_count(), requested.len());
    let commit = edit.commit().unwrap();

    assert_eq!(commit.package().text(), source.text());
    assert_eq!(
        alignments(commit.package()),
        [Alignment::Right, Alignment::Center, Alignment::Left,]
    );
    assert!(
        commit
            .package()
            .body()
            .runs()
            .next()
            .unwrap()
            .format()
            .bold()
    );

    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), source.text());
    assert_eq!(alignments(&reopened), alignments(commit.package()));
}

#[test]
fn batch_shape_bounds_and_conflicts_fail_before_staging() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();

    let mut empty = source.edit();
    assert!(matches!(
        empty.set_body_paragraph_alignments(&[]),
        Err(Error::EmptyParagraphAlignmentBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    for requested in [
        updates(&[(2, Alignment::Center), (1, Alignment::Right)]),
        updates(&[(1, Alignment::Center), (1, Alignment::Right)]),
    ] {
        let mut edit = source.edit();
        assert!(matches!(
            edit.set_body_paragraph_alignments(&requested),
            Err(Error::ParagraphAlignmentBatchOutOfOrder { .. })
        ));
        assert_eq!(edit.operation_count(), 0);
    }

    let late = updates(&[(0, Alignment::Center), (3, Alignment::Right)]);
    let mut edit = source.edit();
    assert!(matches!(
        edit.set_body_paragraph_alignments(&late),
        Err(Error::ParagraphOutOfRange {
            position: 3,
            count: 3
        })
    ));
    assert_eq!(edit.operation_count(), 0);

    let mut duplicate = source.edit();
    duplicate
        .set_paragraph_alignment(Position::new(2), Alignment::Center)
        .unwrap();
    let same = updates(&[(1, Alignment::Right), (2, Alignment::Left)]);
    assert!(matches!(
        duplicate.set_body_paragraph_alignments(&same),
        Err(Error::Conflict {
            existing: 0,
            incoming: 2
        })
    ));
    assert_eq!(duplicate.operation_count(), 1);

    let limits = SemanticLimits::builder()
        .max_operations(1)
        .unwrap()
        .build()
        .unwrap();
    let mut bounded = source.edit_with_options(EditOptions::new().semantic_limits(limits));
    let two = updates(&[(0, Alignment::Center), (1, Alignment::Right)]);
    assert!(matches!(
        bounded.set_body_paragraph_alignments(&two),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 2,
            maximum: 1
        })
    ));
    assert_eq!(bounded.operation_count(), 0);
}

#[test]
fn alignment_batch_composes_with_non_structural_text_and_inverse_is_exact() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let requested = updates(&[(0, Alignment::Center), (2, Alignment::Justify)]);

    let mut edit = source.edit();
    edit.replace_text(TextSpan::new(0, 5).unwrap(), "One")
        .unwrap();
    edit.set_body_paragraph_alignments(&requested).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.diagnostics().operation_count(), 3);
    assert_eq!(commit.package().text(), "One\nSecond\nThird");
    assert_eq!(
        alignments(commit.package()),
        [Alignment::Center, Alignment::Left, Alignment::Justify,]
    );
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn all_noop_alignment_updates_share_the_source_snapshot_and_batches_are_reusable() {
    let source = Package::parse(r"{\rtf1\ansi\qc First\par\qr Second}").unwrap();
    let requested = updates(&[(0, Alignment::Center), (1, Alignment::Right)]);

    let mut first = source.edit();
    first.set_body_paragraph_alignments(&requested).unwrap();
    let first = first.commit().unwrap();
    assert!(!first.diagnostics().changed());
    assert!(first.package().same_snapshot(&source));
    assert!(
        first
            .patch()
            .to_durable(litchi_core::patch::PatchLimits::new(
                litchi_core::patch::BlobLimits::new(0, 0, 0),
                1024 * 1024,
                8,
                8,
                256 * 1024,
                512 * 1024,
            ))
            .unwrap()
            .operations()
            .is_empty()
    );

    let mut second = source.edit();
    second.set_body_paragraph_alignments(&requested).unwrap();
    let second = second.commit().unwrap();
    assert_eq!(
        second.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
    assert_eq!(alignments(second.package()), alignments(first.package()));
}
