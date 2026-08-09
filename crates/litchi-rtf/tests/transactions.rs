#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Alignment, Document,
    edit::{Error, History, HistoryLimits, Limits, TextSpan},
};

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
fn body_text_commit_is_atomic_reversible_and_source_checked() {
    let source = Document::parse(r"{\rtf1\ansi Plain\par Body}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Changed\nbody").unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.text(), "Plain\nBody");
    assert_eq!(commit.snapshot().text(), "Changed\nbody");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 1);

    let applied = commit.patch().apply(&source).unwrap();
    assert!(applied.same_snapshot(commit.snapshot()));
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert!(restored.same_snapshot(&source));

    let other = Document::parse(r"{\rtf1\ansi Other}").unwrap();
    assert!(matches!(
        commit.patch().apply(&other),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn body_text_noop_shares_the_source_snapshot() {
    let source = Document::parse(r"{\rtf1\ansi Same}").unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Same").unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.diagnostics().changed());
    assert!(commit.snapshot().same_snapshot(&source));
}

#[test]
fn changed_body_text_refuses_opaque_syntax_without_touching_the_source() {
    let source = Document::parse(r"{\rtf1\ansi A\future42 B}").unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.replace_body_text("Changed").unwrap();

    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(source.text(), "AB");
}

#[test]
fn transaction_composes_disjoint_spans_and_a_property() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta Gamma}").unwrap();
    let mut edit = source.edit();
    edit.replace_text(TextSpan::new(11, 16).unwrap(), "G")
        .unwrap();
    edit.replace_text(TextSpan::new(0, 5).unwrap(), "A")
        .unwrap();
    edit.set_paragraph_alignment(0, Alignment::Center).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.snapshot().text(), "A Beta G");
    assert_eq!(
        commit
            .snapshot()
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .format()
            .alignment(),
        Alignment::Center
    );
    assert_eq!(commit.diagnostics().operation_count(), 3);
}

#[test]
fn overlapping_spans_duplicate_properties_and_operation_bounds_conflict() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let mut edit = source.edit();
    edit.replace_text(TextSpan::new(0, 5).unwrap(), "A")
        .unwrap();
    assert!(matches!(
        edit.replace_text(TextSpan::new(4, 7).unwrap(), "overlap"),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));

    let mut properties = source.edit();
    properties
        .set_paragraph_alignment(0, Alignment::Center)
        .unwrap();
    assert!(matches!(
        properties.set_paragraph_alignment(0, Alignment::Right),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));

    let mut bounded = source.edit_with_limits(Limits::new(1));
    bounded
        .replace_text(TextSpan::new(0, 1).unwrap(), "a")
        .unwrap();
    assert!(matches!(
        bounded.replace_text(TextSpan::new(2, 3).unwrap(), "p"),
        Err(Error::OperationLimit {
            observed: 2,
            limit: 1
        })
    ));
}

#[test]
fn structural_and_property_operations_fail_closed() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second}").unwrap();
    let mut edit = source.edit();
    edit.set_paragraph_alignment(1, Alignment::Right).unwrap();
    assert!(matches!(
        edit.replace_paragraph_text(0, "One\nInserted"),
        Err(Error::StructuralPropertyConflict)
    ));
}

#[test]
fn body_text_splice_preserves_modeled_header_bytes_exactly() {
    let prefix = br"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss Arial;}}{\colortbl;\red1\green2\blue3;}{\stylesheet{\s0 Normal;}}{\info{\title Preserved;}}\f0 ";
    let source = [prefix.as_slice(), b"Original", b"}"].concat();
    let document = Document::from_bytes(&source).unwrap();
    let mut edit = document.edit();
    edit.replace_body_text("Changed\nbody").unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.snapshot().to_bytes().unwrap();

    assert_eq!(&output[..prefix.len()], prefix);
    assert_eq!(&output[prefix.len()..], br"Changed\par body}");
    assert_eq!(commit.snapshot().text(), "Changed\nbody");
}

#[test]
fn body_splice_preserves_unknown_metadata_group_and_surrounding_writer_bytes() {
    let prefix = br"{\rtf1\ansi{\*\futuremeta vendor-byte-payload}{\info{\title Exact;}}\pard ";
    let source = [prefix.as_slice(), b"Original", b"}\r\n"].concat();
    let document = Document::from_bytes(&source).unwrap();
    let mut edit = document.edit();
    edit.replace_text(TextSpan::new(0, 8).unwrap(), "Changed")
        .unwrap();
    let output = edit.commit().unwrap().snapshot().to_bytes().unwrap();

    assert_eq!(&output[..prefix.len()], prefix);
    assert_eq!(&output[prefix.len()..], b"Changed}\r\n");
}

#[test]
fn paragraph_text_edit_is_checked_local_and_reversible() {
    let source = Document::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(1, "Changed").unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(source.text(), "First\nSecond\nThird");
    assert_eq!(commit.snapshot().text(), "First\nChanged\nThird");
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
    assert!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .same_snapshot(&source)
    );
}

#[test]
fn paragraph_text_edit_rejects_out_of_range_before_staging() {
    let source = Document::parse(r"{\rtf1 Only}").unwrap();
    let mut edit = source.edit();
    assert!(matches!(
        edit.replace_paragraph_text(1, "never"),
        Err(Error::ParagraphOutOfRange {
            position: 1,
            count: 1
        })
    ));
    edit.replace_paragraph_text(0, "Changed").unwrap();
    assert_eq!(edit.commit().unwrap().snapshot().text(), "Changed");
}

#[test]
fn durable_multi_operation_patch_is_deterministic_reversible_and_source_checked() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta Gamma}").unwrap();
    let mut edit = source.edit();
    edit.replace_text(TextSpan::new(11, 16).unwrap(), "G")
        .unwrap();
    edit.replace_text(TextSpan::new(0, 5).unwrap(), "A")
        .unwrap();
    edit.set_paragraph_alignment(0, Alignment::Right).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(3)).unwrap();
    let first = durable.to_deterministic_json().unwrap();
    let second = durable.to_deterministic_json().unwrap();
    assert_eq!(first, second);

    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &first,
            durable_limits(3),
        )
        .unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert_eq!(
        restored
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .format()
            .alignment(),
        Alignment::Left
    );

    let other = Document::parse(r"{\rtf1\ansi Other}").unwrap();
    assert!(matches!(
        other.apply_durable(&decoded),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn durable_patch_reports_a_stale_semantic_precondition() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("text".to_string(), Value::String("stale".to_string()));
    let forward = PatchOperation::new(
        limits,
        "body-text.replace",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::String("Changed".to_string()),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "body-text.replace",
        "body:utf8:0-7",
        preconditions,
        Value::String("Alpha".to_string()),
    )
    .unwrap();
    let patch = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();

    assert!(matches!(
        source.apply_durable(&patch),
        Err(Error::StalePrecondition("body text differs"))
    ));
}

#[test]
fn no_op_patch_is_empty_durable_and_history_is_budgeted() {
    let source = Document::parse(r"{\rtf1\ansi Same}").unwrap();
    let mut edit = source.edit();
    edit.replace_text(TextSpan::new(0, 4).unwrap(), "Same")
        .unwrap();
    edit.set_paragraph_alignment(0, Alignment::Left).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.snapshot().same_snapshot(&source));
    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
    assert!(durable.operations().is_empty());
    assert!(
        source
            .apply_durable(&durable)
            .unwrap()
            .same_snapshot(&source)
    );

    let mut changed = source.edit();
    changed.replace_body_text("Changed").unwrap();
    let changed = changed.commit().unwrap();
    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024));
    history
        .record(changed.snapshot().clone(), changed.patch().history_weight())
        .unwrap();
    assert!(history.undo());
    assert!(history.current().same_snapshot(&source));
    assert!(history.redo());
    assert_eq!(history.current().text(), "Changed");
}
