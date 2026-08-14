#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{Document, ProtectionType, edit::Error};

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        max_operations,
        8,
        256 * 1024,
        512 * 1024,
    )
}

#[test]
fn split_is_lossless_durable_reversible_and_sequential() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let source_bytes = source.to_bytes().unwrap();

    let mut edit = source.edit();
    edit.split_paragraph(1, 3).unwrap();
    let commit = edit.commit().unwrap();
    let expected = br"{\rtf1\ansi First\par Sec\par ond\par Third}";

    assert_eq!(commit.snapshot().text(), "First\nSec\nond\nThird");
    assert_eq!(commit.snapshot().to_bytes().unwrap(), expected);
    let mut streamed = Vec::new();
    commit.snapshot().write_to(&mut streamed).unwrap();
    assert_eq!(streamed, expected);

    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), expected);
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);

    let foreign = Document::parse(r"{\rtf1\ansi Foreign\par Source}").unwrap();
    assert!(matches!(
        foreign.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));

    let leading_empty = Document::parse(r"{\rtf1\ansi First\par}").unwrap();
    let mut leading_edit = leading_empty.edit();
    leading_edit.split_paragraph(0, 0).unwrap();
    let leading_commit = leading_edit.commit().unwrap();
    let leading_durable = leading_commit
        .patch()
        .to_durable(durable_limits(1))
        .unwrap();
    let leading_applied = leading_empty.apply_durable(&leading_durable).unwrap();
    assert_eq!(
        leading_applied
            .apply_durable(&leading_durable.inverse())
            .unwrap()
            .to_bytes()
            .unwrap(),
        leading_empty.to_bytes().unwrap()
    );
}

#[test]
fn durable_split_rejects_forged_noncanonical_boundary() {
    use litchi_core::patch::{BlobBundle, Patch, ReversibleOperation};
    use serde_json::Value;

    let source = Document::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut edit = source.edit();
    edit.split_paragraph(1, 6).unwrap();
    let durable = edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(durable_limits(1))
        .unwrap();

    let mut forged_split = durable.operations()[0].clone();
    forged_split.preconditions.insert(
        "boundary".to_string(),
        Value::String("5c706172".to_string()),
    );
    forged_split.preconditions.insert(
        "boundary_mode".to_string(),
        Value::String("exact".to_string()),
    );
    let forged = Patch::<litchi_core::patch::Reversible>::new(
        durable_limits(1),
        "litchi-rtf",
        [ReversibleOperation::new(
            forged_split,
            durable.inverse().operations()[0].clone(),
        )],
        BlobBundle::new(durable_limits(1).blobs()),
        BlobBundle::new(durable_limits(1).blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&forged),
        Err(Error::StalePrecondition("durable paragraph result differs"))
    ));
}

#[test]
fn merge_removes_only_the_adjacent_boundary() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.merge_paragraphs(0, 1).unwrap();
    let commit = edit.commit().unwrap();
    let expected = br"{\rtf1\ansi FirstSecond\par Third}";

    assert_eq!(commit.snapshot().text(), "FirstSecond\nThird");
    assert_eq!(commit.snapshot().to_bytes().unwrap(), expected);

    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), expected);
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);

    let mut non_adjacent = source.edit();
    assert!(matches!(
        non_adjacent.merge_paragraphs(0, 2),
        Err(Error::ParagraphMergeNonAdjacent {
            first: 0,
            second: 2
        })
    ));
    assert_eq!(non_adjacent.operation_count(), 0);
}

#[test]
fn merge_preserves_exact_boundary_bytes_in_durable_inverse() {
    use litchi_core::patch::{BlobBundle, Patch, ReversibleOperation};
    use serde_json::Value;

    let source = Document::parse(r"{\rtf1\ansi First\par\par}").unwrap();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.merge_paragraphs(0, 1).unwrap();
    let commit = edit.commit().unwrap();
    let expected = br"{\rtf1\ansi First\par}";
    assert_eq!(commit.snapshot().to_bytes().unwrap(), expected);

    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), expected);
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);

    let mut forged_merge = durable.operations()[0].clone();
    forged_merge.preconditions.insert(
        "boundary".to_string(),
        Value::String("5c70617220".to_string()),
    );
    let forged = Patch::<litchi_core::patch::Reversible>::new(
        durable_limits(1),
        "litchi-rtf",
        [ReversibleOperation::new(
            forged_merge,
            durable.inverse().operations()[0].clone(),
        )],
        BlobBundle::new(durable_limits(1).blobs()),
        BlobBundle::new(durable_limits(1).blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&forged),
        Err(Error::StalePrecondition(
            "ordinary paragraph merge boundary bytes differ"
        ))
    ));
}

#[test]
fn split_and_merge_refuse_external_and_dynamic_metadata_exactly() {
    let sources = [
        r"{\rtf1\ansi{\*\xform transform.xsl}First\par Second}",
        r"{\rtf1\ansi\usexform First\par Second}",
        r"{\rtf1\ansi{\*\template template.dot}First\par Second}",
        r"{\rtf1\ansi{\*\mailmerge{\*\mmdatasource datasource.csv}}First\par Second}",
    ];

    for source_text in sources {
        let source = Document::parse(source_text).unwrap();
        let original = source.to_bytes().unwrap();

        let mut split = source.edit();
        assert!(matches!(
            split.split_paragraph(0, 1),
            Err(Error::UnsupportedSource(_))
        ));
        assert_eq!(split.operation_count(), 0);
        assert_eq!(source.to_bytes().unwrap(), original);

        let mut merge = source.edit();
        assert!(matches!(
            merge.merge_paragraphs(0, 1),
            Err(Error::UnsupportedSource(_))
        ));
        assert_eq!(merge.operation_count(), 0);
        assert_eq!(source.to_bytes().unwrap(), original);
    }
}

#[test]
fn split_preflights_boundaries_limits_and_unsafe_sources_atomically() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second}").unwrap();

    let mut out_of_range = source.edit();
    assert!(matches!(
        out_of_range.split_paragraph(0, usize::MAX),
        Err(Error::ParagraphSplitOffsetOutOfRange {
            position: 0,
            offset: usize::MAX,
            length: 5
        })
    ));
    assert_eq!(out_of_range.operation_count(), 0);

    let mut limited = source.edit_with_limits(litchi_rtf::edit::Limits::new(0));
    assert!(matches!(
        limited.split_paragraph(0, 1),
        Err(Error::OperationLimit {
            observed: 1,
            limit: 0
        })
    ));
    assert_eq!(limited.operation_count(), 0);

    let unicode = Document::parse(r"{\rtf1\ansi Caf\'e9}").unwrap();
    let mut non_boundary = unicode.edit();
    assert!(matches!(
        non_boundary.split_paragraph(0, 4),
        Err(Error::SpanNotOnCharacterBoundary { .. })
    ));
    assert_eq!(non_boundary.operation_count(), 0);

    let terminal = Document::parse(r"{\rtf1\ansi First}").unwrap();
    let mut terminal_edit = terminal.edit();
    assert!(matches!(
        terminal_edit.split_paragraph(0, 5),
        Err(Error::ParagraphSplitAtEndRequiresBoundary { position: 0 })
    ));
    assert_eq!(terminal_edit.operation_count(), 0);

    for unsafe_source in [
        r"{\rtf1\ansi First{\b nested}Second}",
        r"{\rtf1\ansi First\qr Second}",
    ] {
        let source = Document::parse(unsafe_source).unwrap();
        let mut edit = source.edit();
        let result = edit.split_paragraph(0, 1);
        assert!(
            result.is_err(),
            "unsafe source unexpectedly accepted: {unsafe_source}"
        );
        assert_eq!(edit.operation_count(), 0);
    }

    let protected =
        Document::parse(r"{\rtf1\ansi\readprot\enforceprot1 First\par Second}").unwrap();
    let mut edit = protected.edit();
    edit.split_paragraph(0, 1).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::ProtectedDocument {
            protection_type: ProtectionType::ReadOnly
        })
    ));

    let undelimited_empty = Document::parse(r"{\rtf1\ansi\par First}").unwrap();
    let mut undelimited_edit = undelimited_empty.edit();
    assert!(matches!(
        undelimited_edit.merge_paragraphs(0, 1),
        Err(Error::UnsupportedSource(_))
    ));
    assert_eq!(undelimited_edit.operation_count(), 0);
}
