#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Alignment, CharacterBaseline, Document, HeaderFooterType, TableCellPath, UnderlineStyle,
    edit::{
        Composition, CompositionError, CompositionLimits, Error, HeaderFooterParagraph, History,
        HistoryLimits, Limits, MergePlan, MergeResolution, TextSpan, TransferPlan,
    },
};
use std::num::NonZeroU16;

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
    let mut overlap = source.edit();
    overlap
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    assert!(matches!(
        overlap.set_text_underline(TextSpan::new(4, 5).unwrap(), UnderlineStyle::Double),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    let mut empty_edit = source.edit();
    assert!(matches!(
        empty_edit.set_text_underline(TextSpan::new(0, 0).unwrap(), UnderlineStyle::Single),
        Err(Error::UnsupportedSource(
            "underline edits require non-empty text within one paragraph"
        ))
    ));
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

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024));
    let mut changed = history.current().edit();
    changed.replace_body_text("Changed").unwrap();
    history.commit(changed).unwrap();
    assert!(history.undo());
    assert!(history.current().same_snapshot(&source));
    assert!(history.redo());
    assert_eq!(history.current().text(), "Changed");
}

#[test]
fn bold_property_and_paragraph_structure_have_durable_semantics() {
    let probe = Document::parse(r"{\rtf1\ansi \ql \b Alpha\b0 \par \ql Beta}").unwrap();
    assert!(probe.body().runs().next().unwrap().format().bold());
    let source = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut formatting = source.edit();
    formatting
        .set_text_bold(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let formatted = formatting.commit().unwrap();
    let runs = formatted.snapshot().body().runs().collect::<Vec<_>>();
    assert_eq!(runs[0].text(), "Alpha");
    assert!(runs[0].format().bold());
    assert!(!runs[1].format().bold());

    let durable = formatted.patch().to_durable(durable_limits(1)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert!(applied.body().runs().next().unwrap().format().bold());
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(!restored.body().runs().next().unwrap().format().bold());

    let mut structure = source.edit();
    structure.insert_paragraph_after(0, "Inserted").unwrap();
    let inserted = structure.commit().unwrap();
    assert_eq!(inserted.snapshot().text(), "Alpha\nInserted\nBeta");
    let structural_patch = inserted.patch().to_durable(durable_limits(1)).unwrap();
    let structurally_applied = source.apply_durable(&structural_patch).unwrap();
    assert_eq!(structurally_applied.text(), "Alpha\nInserted\nBeta");
    assert_eq!(
        structurally_applied
            .apply_durable(&structural_patch.inverse())
            .unwrap()
            .text(),
        source.text()
    );
}

#[test]
fn multiple_paragraph_insertions_have_target_relative_durable_inverses() {
    let source = Document::parse(r"{\rtf1\ansi A\par B\par C}").unwrap();
    let mut edit = source.edit();
    edit.insert_paragraph_after(0, "X")
        .unwrap()
        .insert_paragraph_after(1, "Y")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().text(), "A\nX\nB\nY\nC");

    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
    let durable_json: serde_json::Value =
        serde_json::from_slice(&durable.to_deterministic_json().unwrap()).unwrap();
    let targets = durable_json["operations"]
        .as_array()
        .unwrap()
        .iter()
        .map(|operation| operation["forward"]["target"].as_str().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(targets, ["body:paragraph:0", "body:paragraph:1"]);

    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );

    let inverse = durable.inverse();
    let inverse_json: serde_json::Value =
        serde_json::from_slice(&inverse.to_deterministic_json().unwrap()).unwrap();
    let targets = inverse_json["operations"]
        .as_array()
        .unwrap()
        .iter()
        .map(|operation| operation["forward"]["target"].as_str().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(targets, ["body:paragraph:2", "body:paragraph:0"]);

    let restored = applied.apply_durable(&inverse).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn italic_property_is_bounded_batch_reversible_and_durable() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta Gamma}").unwrap();
    let mut edit = source.edit();
    edit.set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap()
        .set_text_italic(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let formatted = edit.commit().unwrap();
    let runs = formatted.snapshot().body().runs().collect::<Vec<_>>();
    assert_eq!(
        runs.iter()
            .map(|run| (run.text(), run.format().italic()))
            .collect::<Vec<_>>(),
        vec![
            ("Alpha", true),
            (" ", false),
            ("Beta", true),
            (" Gamma", false)
        ]
    );
    assert_eq!(formatted.snapshot().text(), source.text());
    assert_eq!(
        formatted
            .patch()
            .inverse()
            .apply(formatted.snapshot())
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    let prefix = br"{\rtf1\ansi{\*\futuremeta retained}\pard ";
    let enveloped =
        Document::from_bytes(&[prefix.as_slice(), b"Alpha", b"}\r\n"].concat()).unwrap();
    let mut enveloped_edit = enveloped.edit();
    enveloped_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let enveloped_bytes = enveloped_edit
        .commit()
        .unwrap()
        .snapshot()
        .to_bytes()
        .unwrap();
    assert!(enveloped_bytes.starts_with(prefix));
    assert!(enveloped_bytes.ends_with(b"}\r\n"));

    let durable = formatted.patch().to_durable(durable_limits(2)).unwrap();
    assert_eq!(durable.operations().len(), 2);
    assert!(
        durable
            .operations()
            .iter()
            .all(|operation| operation.op == "character-italic.set")
    );
    let encoded = durable.to_deterministic_json().unwrap();
    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &encoded,
            durable_limits(2),
        )
        .unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        formatted.snapshot().to_bytes().unwrap()
    );
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| !run.format().italic()));

    let mut mixed_edit = source.edit();
    mixed_edit
        .set_text_bold(TextSpan::new(0, 5).unwrap(), true)
        .unwrap()
        .set_text_italic(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let mixed = mixed_edit.commit().unwrap();
    let mixed_runs = mixed.snapshot().body().runs().collect::<Vec<_>>();
    assert!(
        mixed_runs
            .iter()
            .any(|run| run.text() == "Alpha" && run.format().bold())
    );
    assert!(
        mixed_runs
            .iter()
            .any(|run| run.text() == "Beta" && run.format().italic())
    );
}

#[test]
fn italic_property_noop_conflict_and_source_closure_refusals_are_atomic() {
    let source = Document::parse(r"{\rtf1\ansi {\i Alpha} Beta}").unwrap();
    let original = source.to_bytes().unwrap();
    let mut noop = source.edit();
    noop.set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.snapshot().to_bytes().unwrap(), original);

    let plain = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let mut overlap = plain.edit();
    overlap
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        overlap.set_text_italic(TextSpan::new(4, 10).unwrap(), true),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    let mut bounded = plain.edit_with_limits(Limits::new(1));
    bounded
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        bounded.set_text_italic(TextSpan::new(6, 10).unwrap(), true),
        Err(Error::OperationLimit {
            observed: 2,
            limit: 1
        })
    ));

    let utf8 = Document::parse(r"{\rtf1\ansi Caf\'e9}").unwrap();
    assert_eq!(utf8.text(), "Café");
    let mut utf8_edit = utf8.edit();
    utf8_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(utf8_edit.commit().is_ok());
    let mut misaligned = utf8.edit();
    assert!(matches!(
        misaligned.set_text_italic(TextSpan::new(4, 5).unwrap(), true),
        Err(Error::SpanNotOnCharacterBoundary { position: 4 })
    ));

    let mixed = Document::parse(r"{\rtf1\ansi {\i Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_italic(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed italic state"
        ))
    ));

    let unrelated = Document::parse(r"{\rtf1\ansi \b Alpha\b0 Beta}").unwrap();
    let mut unrelated_edit = unrelated.edit();
    unrelated_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let unrelated_result = unrelated_edit.commit();
    assert!(matches!(
        unrelated_result,
        Err(Error::UnsupportedSource(
            "the body has mixed character formatting"
        ))
    ));

    let protected = Document::parse(r"{\rtf1\ansi\readprot\enforceprot1 Alpha}").unwrap();
    let mut protected_edit = protected.edit();
    protected_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        protected_edit.commit(),
        Err(Error::ProtectedDocument { .. })
    ));

    let paragraph_source = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph_source.edit();
    assert!(matches!(
        paragraph_edit.set_text_italic(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "italic edits require non-empty text within one paragraph"
        ))
    ));

    let opaque = Document::parse(r"{\rtf1\ansi Alpha{\future42 retained}}").unwrap();
    let mut opaque_edit = opaque.edit();
    opaque_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        opaque_edit.commit(),
        Err(Error::UnsupportedSource(_))
    ));

    let field = Document::parse(r"{\rtf1\ansi A{\field{\*\fldinst PAGE}{\fldrslt B}}}").unwrap();
    let mut field_edit = field.edit();
    field_edit
        .set_text_italic(TextSpan::new(0, 1).unwrap(), true)
        .unwrap();
    assert!(matches!(
        field_edit.commit(),
        Err(Error::UnsupportedSource(_))
    ));

    let table = Document::parse(r"{\rtf1\ansi\trowd\cellx1000\intbl A\cell\row}").unwrap();
    let mut table_edit = table.edit();
    assert!(matches!(
        table_edit.set_text_italic(TextSpan::new(0, 1).unwrap(), true),
        Err(Error::SpanOutOfRange { .. })
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    assert_eq!(cp1252.text(), "Café");
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "italic edits refuse non-ASCII transport encodings"
        ))
    ));

    let compressed_bytes = litchi_rtf::transport::compress(br"{\rtf1\ansi Alpha}", true).unwrap();
    let compressed = Document::from_bytes(&compressed_bytes).unwrap();
    let mut compressed_edit = compressed.edit();
    compressed_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        compressed_edit.commit(),
        Err(Error::UnsupportedSource(
            "compressed RTF needs a transport-aware rewrite"
        ))
    ));
}

#[test]
fn underline_edit_preserves_uniform_body_character_baseline_and_inverse() {
    let source = Document::parse(r"{\rtf1\ansi\b\uldb Alpha Beta}").unwrap();
    let mut edit = source.edit();
    edit.set_text_underline(TextSpan::new(6, 10).unwrap(), UnderlineStyle::Single)
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
    assert!(reopened.body().runs().all(|run| run.format().bold()));
    assert_eq!(
        reopened
            .body()
            .runs()
            .find(|run| run.text() == "Alpha ")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::Double
    );
    assert_eq!(
        reopened
            .body()
            .runs()
            .find(|run| run.text() == "Beta")
            .unwrap()
            .format()
            .underline(),
        UnderlineStyle::Single
    );
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored.to_bytes().unwrap(),
        source.to_bytes().unwrap(),
        "inverse must restore the exact source artifact"
    );
}

#[test]
fn underline_property_supports_exact_styles_and_inverse() {
    let styles = [
        UnderlineStyle::None,
        UnderlineStyle::Single,
        UnderlineStyle::Double,
        UnderlineStyle::Dotted,
        UnderlineStyle::Dashed,
        UnderlineStyle::DashDot,
        UnderlineStyle::DashDotDot,
        UnderlineStyle::Words,
        UnderlineStyle::Thick,
        UnderlineStyle::Wave,
        UnderlineStyle::Hairline,
        UnderlineStyle::ThickDotted,
        UnderlineStyle::ThickDashed,
        UnderlineStyle::ThickDashDot,
        UnderlineStyle::ThickDashDotDot,
        UnderlineStyle::ThickLongDash,
        UnderlineStyle::LongDash,
        UnderlineStyle::HeavyWave,
        UnderlineStyle::DoubleWave,
    ];
    for style in styles {
        let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
        let mut edit = source.edit();
        edit.set_text_underline(TextSpan::new(0, 5).unwrap(), style)
            .unwrap();
        let commit = edit.commit().unwrap();
        if style == UnderlineStyle::None {
            assert!(!commit.diagnostics().changed());
        }
        assert_eq!(
            commit
                .snapshot()
                .body()
                .runs()
                .next()
                .unwrap()
                .format()
                .underline(),
            style
        );
        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert!(
            restored
                .body()
                .runs()
                .all(|run| { run.format().underline() == UnderlineStyle::None })
        );
    }
}

#[test]
fn underline_property_refuses_mixed_structure_transport_and_stale_durable_state() {
    let mixed = Document::parse(r"{\rtf1\ansi {\ul Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_underline(TextSpan::new(0, 10).unwrap(), UnderlineStyle::None),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed underline state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_underline(TextSpan::new(0, 6).unwrap(), UnderlineStyle::Single),
        Err(Error::UnsupportedSource(
            "underline edits require non-empty text within one paragraph"
        ))
    ));

    let opaque = Document::parse(r"{\rtf1\ansi Alpha{\future42 retained}}").unwrap();
    let mut opaque_edit = opaque.edit();
    opaque_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    assert!(matches!(
        opaque_edit.commit(),
        Err(Error::UnsupportedSource(_))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut misaligned = cp1252.edit();
    assert!(matches!(
        misaligned.set_text_underline(TextSpan::new(4, 5).unwrap(), UnderlineStyle::Single),
        Err(Error::SpanNotOnCharacterBoundary { position: 4 })
    ));
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "underline edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut edit = source.edit();
    edit.set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Double)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-underline.set");
    assert_eq!(
        durable.operations()[0].preconditions["underline"],
        Value::String("none".to_string())
    );
    assert_eq!(
        durable.operations()[0].value,
        Value::String("double".to_string())
    );
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(
        applied.body().runs().next().unwrap().format().underline(),
        UnderlineStyle::Double
    );
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(
        restored
            .body()
            .runs()
            .all(|run| { run.format().underline() == UnderlineStyle::None })
    );

    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("underline".to_string(), Value::String("single".to_string()));
    let forward = PatchOperation::new(
        limits,
        "character-underline.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::String("double".to_string()),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-underline.set",
        "body:utf8:0-5",
        preconditions,
        Value::String("single".to_string()),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition(
            "character underline state differs"
        ))
    ));

    let mut valid_preconditions = BTreeMap::new();
    valid_preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    valid_preconditions.insert("underline".to_string(), Value::String("none".to_string()));
    let invalid_forward = PatchOperation::new(
        limits,
        "character-underline.set",
        "body:utf8:0-5",
        valid_preconditions.clone(),
        Value::String("invalid".to_string()),
    )
    .unwrap();
    let invalid_inverse = PatchOperation::new(
        limits,
        "character-underline.set",
        "body:utf8:0-5",
        valid_preconditions,
        Value::String("none".to_string()),
    )
    .unwrap();
    let invalid = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(invalid_forward, invalid_inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&invalid),
        Err(Error::DurablePatch(message))
            if message == "underline value must be a string"
    ));
}

#[test]
fn body_opaque_validation_cannot_be_masked_by_duplicate_metadata_bytes() {
    let bytes = br"{\rtf1\ansi{\*\future42 retained}Alpha{\future42 retained}}";
    let source = Document::from_bytes(bytes).unwrap();
    let mut edit = source.edit();
    edit.set_text_underline(TextSpan::new(0, 5).unwrap(), UnderlineStyle::Single)
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
}

#[test]
fn strike_edit_preserves_double_strike_and_creates_later_run_boundary() {
    let source = Document::parse(r"{\rtf1\ansi\striked First Second}").unwrap();
    let mut edit = source.edit();
    edit.set_text_strike(TextSpan::new(6, 12).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\striked".len())
            .any(|window| window == br"\striked")
    );
    assert!(
        bytes
            .windows(br"\strike ".len())
            .any(|window| window == br"\strike ")
    );
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert!(
        reopened
            .body()
            .runs()
            .all(|run| run.format().double_strike())
    );
    assert_eq!(
        reopened
            .body()
            .runs()
            .find(|run| run.text() == "Second")
            .unwrap()
            .format()
            .strike(),
        true
    );
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
}

#[test]
fn strike_property_refuses_mixed_structure_transport_and_stale_durable_state() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let mixed = Document::parse(r"{\rtf1\ansi {\strike Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_strike(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed strike state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_strike(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "strike edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "strike edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut edit = source.edit();
    edit.set_text_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-strike.set");
    assert_eq!(
        durable.operations()[0].preconditions["strike"],
        Value::Bool(false)
    );
    assert_eq!(durable.operations()[0].value, Value::Bool(true));
    let applied = source.apply_durable(&durable).unwrap();
    assert!(applied.body().runs().all(|run| run.format().strike()));
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().strike()));
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("strike".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-strike.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-strike.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition("character strike state differs"))
    ));
}

#[test]
fn hidden_edit_is_noop_when_state_is_already_hidden_or_visible() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let mut visible = source.edit();
    visible
        .set_text_hidden(TextSpan::new(0, 5).unwrap(), false)
        .unwrap();
    let visible_commit = visible.commit().unwrap();
    assert!(!visible_commit.diagnostics().changed());
    assert!(visible_commit.snapshot().same_snapshot(&source));

    let mut hidden = source.edit();
    hidden
        .set_text_hidden(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let hidden_commit = hidden.commit().unwrap();
    let bytes = hidden_commit.snapshot().to_bytes().unwrap();
    assert!(bytes.windows(br"\v ".len()).any(|window| window == br"\v "));
    assert!(
        bytes
            .windows(br"\v0 ".len())
            .any(|window| window == br"\v0 ")
    );
    let reopened = Document::from_bytes(&bytes).unwrap();
    let hidden_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().hidden())
        .map(|run| run.text())
        .collect::<String>();
    let visible_text = reopened
        .body()
        .runs()
        .filter(|run| !run.format().hidden())
        .map(|run| run.text())
        .collect::<String>();
    assert!(hidden_text.contains("Alpha"));
    assert!(visible_text.contains("Beta"));
    let restored = hidden_commit
        .patch()
        .inverse()
        .apply(hidden_commit.snapshot())
        .unwrap();
    assert!(restored.body().runs().all(|run| !run.format().hidden()));
}

#[test]
fn hidden_property_preserves_run_boundaries_and_refuses_mixed_structure_transport_and_stale_state()
{
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let mut edit = source.edit();
    edit.set_text_hidden(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    let reopened = Document::from_bytes(&bytes).unwrap();
    let hidden_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().hidden())
        .map(|run| run.text())
        .collect::<String>();
    let visible_text = reopened
        .body()
        .runs()
        .filter(|run| !run.format().hidden())
        .map(|run| run.text())
        .collect::<String>();
    assert!(hidden_text.contains("Beta"));
    assert!(visible_text.contains("Alpha"));
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().hidden()));

    let mixed = Document::parse(r"{\rtf1\ansi {\v Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_hidden(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed hidden state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_hidden(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "hidden edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_hidden(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "hidden edits refuse non-ASCII transport encodings"
        ))
    ));

    let limits = durable_limits(1);
    let mut durable_edit = source.edit();
    durable_edit
        .set_text_hidden(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let durable_commit = durable_edit.commit().unwrap();
    let durable = durable_commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-hidden.set");
    assert_eq!(
        durable.operations()[0].preconditions["hidden"],
        Value::Bool(false)
    );
    assert_eq!(durable.operations()[0].value, Value::Bool(true));
    let applied = source.apply_durable(&durable).unwrap();
    assert!(applied.body().runs().any(|run| run.format().hidden()));
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().hidden()));
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("hidden".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-hidden.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-hidden.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition("character hidden state differs"))
    ));
}

#[test]
fn small_caps_property_preserves_all_caps_and_run_boundaries() {
    let source = Document::parse(r"{\rtf1\ansi\caps Alpha Beta}").unwrap();
    let mut noop = source.edit();
    noop.set_text_small_caps(TextSpan::new(0, 5).unwrap(), false)
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.snapshot().same_snapshot(&source));

    let mut edit = source.edit();
    edit.set_text_small_caps(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\scaps ".len())
            .any(|window| window == br"\scaps ")
    );
    assert!(
        bytes
            .windows(br"\scaps0 ".len())
            .any(|window| window == br"\scaps0 ")
    );
    let reopened = Document::from_bytes(&bytes).unwrap();
    let small_caps_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().small_caps())
        .map(|run| run.text())
        .collect::<String>();
    let ordinary_text = reopened
        .body()
        .runs()
        .filter(|run| !run.format().small_caps())
        .map(|run| run.text())
        .collect::<String>();
    assert!(ordinary_text.contains("Alpha"));
    assert!(small_caps_text.contains("Beta"));
    assert!(reopened.body().runs().all(|run| run.format().all_caps()));

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().small_caps()));
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
}

#[test]
fn small_caps_property_refuses_mixed_structure_transport_and_stale_durable_state() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let associated = Document::parse(r"{\rtf1\ansi\ascaps Alpha}").unwrap();
    assert!(
        associated
            .body()
            .runs()
            .all(|run| !run.format().small_caps())
    );

    let mixed = Document::parse(r"{\rtf1\ansi {\scaps Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_small_caps(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed small-caps state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_small_caps(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "small-caps edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_small_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "small-caps edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut edit = source.edit();
    edit.set_text_small_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-small-caps.set");
    assert_eq!(
        durable.operations()[0].preconditions["small_caps"],
        Value::Bool(false)
    );
    assert_eq!(durable.operations()[0].value, Value::Bool(true));
    let applied = source.apply_durable(&durable).unwrap();
    assert!(applied.body().runs().any(|run| run.format().small_caps()));
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().small_caps()));
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("small_caps".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-small-caps.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-small-caps.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition(
            "character small-caps state differs"
        ))
    ));
}

#[test]
fn all_caps_property_preserves_small_caps_and_both_control_boundaries() {
    let source = Document::parse(r"{\rtf1\ansi\scaps Alpha Beta}").unwrap();
    let mut noop = source.edit();
    noop.set_text_all_caps(TextSpan::new(0, 5).unwrap(), false)
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.snapshot().same_snapshot(&source));

    let mut edit = source.edit();
    edit.set_text_all_caps(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\caps ".len())
            .any(|window| window == br"\caps ")
    );
    assert!(
        bytes
            .windows(br"\caps0 ".len())
            .any(|window| window == br"\caps0 ")
    );

    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), source.text());
    let small_caps_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().small_caps())
        .map(|run| run.text())
        .collect::<String>();
    let all_caps_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().all_caps())
        .map(|run| run.text())
        .collect::<String>();
    assert!(small_caps_text.contains("Alpha"));
    assert!(small_caps_text.contains("Beta"));
    assert!(all_caps_text.contains("Beta"));
    assert!(!all_caps_text.contains("Alpha"));

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| run.format().small_caps()));
    assert!(restored.body().runs().all(|run| !run.format().all_caps()));
    let reopened_restored = Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened_restored.text(), source.text());
}

#[test]
fn all_caps_property_refuses_mixed_structure_transport_and_stale_durable_state() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let associated = Document::parse(r"{\rtf1\ansi\acaps Alpha}").unwrap();
    assert!(associated.body().runs().all(|run| !run.format().all_caps()));

    let mixed = Document::parse(r"{\rtf1\ansi {\caps Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_all_caps(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed all-caps state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_all_caps(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "all-caps edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_all_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "all-caps edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut edit = source.edit();
    edit.set_text_all_caps(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-all-caps.set");
    assert_eq!(
        durable.operations()[0].preconditions["all_caps"],
        Value::Bool(false)
    );
    assert_eq!(durable.operations()[0].value, Value::Bool(true));
    let applied = source.apply_durable(&durable).unwrap();
    assert!(applied.body().runs().any(|run| run.format().all_caps()));
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().all_caps()));
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("all_caps".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-all-caps.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-all-caps.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition("character all-caps state differs"))
    ));
}

#[test]
fn double_strike_property_preserves_single_strike_and_both_control_boundaries() {
    let source = Document::parse(r"{\rtf1\ansi\strike Alpha Beta}").unwrap();
    let mut noop = source.edit();
    noop.set_text_double_strike(TextSpan::new(0, 5).unwrap(), false)
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.snapshot().same_snapshot(&source));

    let mut edit = source.edit();
    edit.set_text_double_strike(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\striked ".len())
            .any(|window| window == br"\striked ")
    );
    assert!(
        bytes
            .windows(br"\striked0 ".len())
            .any(|window| window == br"\striked0 ")
    );

    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), source.text());
    let single_strike_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().strike())
        .map(|run| run.text())
        .collect::<String>();
    let double_strike_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().double_strike())
        .map(|run| run.text())
        .collect::<String>();
    assert!(single_strike_text.contains("Alpha"));
    assert!(single_strike_text.contains("Beta"));
    assert!(double_strike_text.contains("Beta"));
    assert!(!double_strike_text.contains("Alpha"));

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(restored.body().runs().all(|run| run.format().strike()));
    assert!(
        restored
            .body()
            .runs()
            .all(|run| !run.format().double_strike())
    );
    let reopened_restored = Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened_restored.text(), source.text());
}

#[test]
fn double_strike_property_refuses_mixed_structure_transport_and_stale_durable_state() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::Value;
    use std::collections::BTreeMap;

    let mixed = Document::parse(r"{\rtf1\ansi {\striked Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_double_strike(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed double-strike state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_double_strike(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "double-strike edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_double_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "double-strike edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let limits = durable_limits(1);
    let mut edit = source.edit();
    edit.set_text_double_strike(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(limits).unwrap();
    assert_eq!(durable.operations()[0].op, "character-double-strike.set");
    assert_eq!(
        durable.operations()[0].preconditions["double_strike"],
        Value::Bool(false)
    );
    assert_eq!(durable.operations()[0].value, Value::Bool(true));
    let applied = source.apply_durable(&durable).unwrap();
    assert!(
        applied
            .body()
            .runs()
            .any(|run| run.format().double_strike())
    );
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(
        restored
            .body()
            .runs()
            .all(|run| !run.format().double_strike())
    );
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        Value::String(litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex()),
    );
    preconditions.insert("double_strike".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-double-strike.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-double-strike.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition(
            "character double-strike state differs"
        ))
    ));
}

#[test]
fn italic_durable_patch_rejects_stale_property_and_artifact() {
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
    preconditions.insert("italic".to_string(), Value::Bool(true));
    let forward = PatchOperation::new(
        limits,
        "character-italic.set",
        "body:utf8:0-5",
        preconditions.clone(),
        Value::Bool(false),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "character-italic.set",
        "body:utf8:0-5",
        preconditions,
        Value::Bool(true),
    )
    .unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&stale),
        Err(Error::StalePrecondition("character italic state differs"))
    ));

    let foreign = Document::parse(r"{\rtf1\ansi Other}").unwrap();
    let mut valid_edit = source.edit();
    valid_edit
        .set_text_italic(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let valid = valid_edit.commit().unwrap();
    let durable = valid.patch().to_durable(limits).unwrap();
    assert!(matches!(
        foreign.apply_durable(&durable),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn core_subedits_compose_and_report_typed_conflicts() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);

    let mut text_edit = source.edit();
    text_edit
        .replace_text(TextSpan::new(0, 5).unwrap(), "A")
        .unwrap();
    let prepared_text = text_edit.into_sub_edit("text", limits).unwrap();
    let mut alignment_edit = source.edit();
    alignment_edit
        .set_paragraph_alignment(0, Alignment::Center)
        .unwrap();
    let prepared_alignment = alignment_edit.into_sub_edit("alignment", limits).unwrap();
    let mut joined = Composition::new(&source, limits);
    joined
        .join(prepared_text)
        .unwrap()
        .join(prepared_alignment)
        .unwrap();
    let committed = joined.commit().unwrap();
    assert_eq!(committed.snapshot().text(), "A Beta");
    assert_eq!(
        committed
            .snapshot()
            .body()
            .paragraphs()
            .next()
            .unwrap()
            .format()
            .alignment(),
        Alignment::Center
    );

    let mut left = source.edit();
    left.replace_text(TextSpan::new(0, 5).unwrap(), "Left")
        .unwrap();
    let mut right = source.edit();
    right
        .replace_text(TextSpan::new(6, 10).unwrap(), "Right")
        .unwrap();
    let mut conflicts = Composition::new(&source, limits);
    conflicts
        .join(left.into_sub_edit("left", limits).unwrap())
        .unwrap();
    let error = conflicts
        .join(right.into_sub_edit("right", limits).unwrap())
        .unwrap_err();
    assert!(matches!(error, CompositionError::Conflicts(set) if !set.is_empty()));
}

#[test]
fn three_way_merge_is_non_mutating_until_resolved_and_committed() {
    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut left_edit = source.edit();
    left_edit
        .replace_text(TextSpan::new(0, 5).unwrap(), "Left")
        .unwrap();
    let mut right_edit = source.edit();
    right_edit
        .replace_text(TextSpan::new(6, 10).unwrap(), "Right")
        .unwrap();
    let mut left = Composition::new(&source, limits);
    left.join(left_edit.into_sub_edit("left", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&source, limits);
    right
        .join(right_edit.into_sub_edit("right", limits).unwrap())
        .unwrap();

    let plan = MergePlan::new(left, right).unwrap();
    assert_eq!(source.text(), "Alpha Beta");
    assert!(!plan.conflicts().is_empty());
    let mut unresolved = *plan.finish().unwrap_err();
    unresolved.resolve(MergeResolution::Left);
    let merged = unresolved.finish().unwrap().commit().unwrap();
    assert_eq!(merged.snapshot().text(), "Left Beta");
    assert_eq!(source.text(), "Alpha Beta");
}

#[test]
fn retained_destinations_are_multi_operation_durable_and_reversible() {
    let source = Document::parse(
        r"{\rtf1\ansi{\header Head}\pard Body\par\trowd\cellx1000\cellx2000\intbl A\cell B\cell\row}",
    )
    .unwrap();
    let header = HeaderFooterParagraph::new(0, HeaderFooterType::Header, 0);
    let first_cell = TableCellPath::outer(0, 0, 0);
    let second_cell = TableCellPath::outer(0, 0, 1);
    let mut edit = source.edit();
    edit.set_header_footer_text(header, "Running head").unwrap();
    edit.set_table_cell_text(first_cell.clone(), "First")
        .unwrap();
    edit.set_table_cell_text(second_cell, "Second").unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(
        commit.snapshot().sections()[0].headers_footers[0].text(),
        "Running head"
    );
    assert_eq!(
        commit.snapshot().tables()[0].rows()[0].cells()[0].text(),
        "First"
    );
    assert_eq!(
        commit.snapshot().tables()[0].rows()[0].cells()[1].text(),
        "Second"
    );
    assert_eq!(source.tables()[0].rows()[0].cells()[0].text(), "A");

    let durable = commit.patch().to_durable(durable_limits(3)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.tables()[0].rows()[0].cells()[0].text(), "A");
    assert_eq!(restored.sections()[0].headers_footers[0].text(), "Head");
}

#[test]
fn note_and_annotation_text_are_durable_historical_mergeable_and_reopen() {
    let source = Document::parse(
        r"{\rtf1\ansi A{\footnote\chftn Old note}B{\*\atnid AM}{\*\atnauthor Ada}\chatn{\*\annotation Old comment}C}",
    )
    .unwrap();
    let mut edit = source.edit();
    edit.set_note_text(0, "Updated note").unwrap();
    edit.set_annotation_text(0, "Updated comment").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().notes()[0].content, "Updated note");
    assert_eq!(commit.snapshot().annotations()[0].text, "Updated comment");
    assert_eq!(commit.snapshot().text(), source.text());

    let bytes = commit.snapshot().to_bytes().unwrap();
    let serialized = String::from_utf8_lossy(&bytes);
    assert!(serialized.contains("\\footnote"));
    assert!(serialized.contains("\\annotation"));
    assert!(serialized.contains("\\atnauthor"));
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.notes()[0].content, "Updated note");
    assert!(reopened.notes()[0].is_footnote);
    assert_eq!(reopened.annotations()[0].text, "Updated comment");
    assert_eq!(reopened.annotations()[0].author, "Ada");
    assert!(!reopened.annotations()[0].has_reference);

    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), bytes);
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.notes()[0].content, "Old note");
    assert_eq!(restored.annotations()[0].text, "Old comment");

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024 * 1024));
    history.record_commit(&commit).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().annotations()[0].text, "Old comment");
    assert!(history.redo());
    assert_eq!(history.current().notes()[0].content, "Updated note");

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut note_edit = source.edit();
    note_edit.set_note_text(0, "Merged note").unwrap();
    let mut annotation_edit = source.edit();
    annotation_edit
        .set_annotation_text(0, "Merged comment")
        .unwrap();
    let mut left = Composition::new(&source, limits);
    left.join(note_edit.into_sub_edit("note", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&source, limits);
    right
        .join(annotation_edit.into_sub_edit("annotation", limits).unwrap())
        .unwrap();
    let merge = MergePlan::new(left, right).unwrap();
    assert!(merge.conflicts().is_empty());
    let merged = merge.finish().unwrap().commit().unwrap();
    assert_eq!(merged.snapshot().notes()[0].content, "Merged note");
    assert_eq!(merged.snapshot().annotations()[0].text, "Merged comment");

    let mut first_note = source.edit();
    first_note.set_note_text(0, "First branch").unwrap();
    let mut second_note = source.edit();
    second_note.set_note_text(0, "Second branch").unwrap();
    let mut left = Composition::new(&source, limits);
    left.join(first_note.into_sub_edit("first-note", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&source, limits);
    right
        .join(second_note.into_sub_edit("second-note", limits).unwrap())
        .unwrap();
    assert_eq!(MergePlan::new(left, right).unwrap().conflicts().len(), 1);
}

#[test]
fn note_and_annotation_edits_refuse_positioned_or_opaque_dependencies() {
    let active_note = Document::parse(
        r"{\rtf1 A{\footnote\chftn before{\field{\*\fldinst INCLUDETEXT external}{\fldrslt cached}}after}B}",
    )
    .unwrap();
    let mut active_edit = active_note.edit();
    assert!(matches!(
        active_edit.set_note_text(0, "replacement"),
        Err(Error::UnsupportedSource(_))
    ));

    let opaque_annotation = Document::parse(
        r"{\rtf1{\*\vendor retained}A{\*\atnid AM}{\*\atnauthor Ada}\chatn{\*\annotation Old comment}B}",
    )
    .unwrap();
    let exact = opaque_annotation.to_bytes().unwrap();
    let mut edit = opaque_annotation.edit();
    edit.set_annotation_text(0, "replacement").unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(opaque_annotation.to_bytes().unwrap(), exact);
}

#[test]
fn genuine_libreoffice_shape_text_is_durable_mergeable_transferable_and_reopens() {
    let source = Document::from_bytes(include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf"
    ))
    .unwrap();
    let original = source.shapes()[0].clone();
    let mut edit = source.edit();
    edit.set_shape_text(0, "Edited frame\nsecond line").unwrap();
    let commit = edit.commit().unwrap();
    let edited = &commit.snapshot().shapes()[0];
    assert_eq!(edited.text, "Edited frame\nsecond line");
    assert_eq!(edited.position, original.position);
    assert_eq!(edited.geometry, original.geometry);
    assert_eq!(edited.properties, original.properties);
    assert_eq!(edited.text_formatting, original.text_formatting);
    assert_eq!(commit.snapshot().text(), source.text());

    let output = commit.snapshot().to_bytes().unwrap();
    assert!(String::from_utf8_lossy(&output).contains("\\shptxt Edited frame\\par second line"));
    let reopened = Document::from_bytes(&output).unwrap();
    assert_eq!(reopened.shapes()[0].text, "Edited frame\nsecond line");
    assert_eq!(reopened.shapes()[0].geometry, original.geometry);

    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), output);
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert_eq!(restored.shapes()[0].text, original.text);

    let mut history = History::new(source.clone(), HistoryLimits::new(2, 1024 * 1024));
    history.record_commit(&commit).unwrap();
    assert!(history.undo());
    assert_eq!(history.current().shapes()[0].text, original.text);
    assert!(history.redo());
    assert_eq!(
        history.current().shapes()[0].text,
        "Edited frame\nsecond line"
    );

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut first = source.edit();
    first.set_shape_text(0, "First branch").unwrap();
    let mut second = source.edit();
    second.set_shape_text(0, "Second branch").unwrap();
    let mut left = Composition::new(&source, limits);
    left.join(first.into_sub_edit("first-shape", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&source, limits);
    right
        .join(second.into_sub_edit("second-shape", limits).unwrap())
        .unwrap();
    assert_eq!(MergePlan::new(left, right).unwrap().conflicts().len(), 1);

    let target = Document::parse(r"{\rtf1 Target}").unwrap();
    let transfer = TransferPlan::shape(&source, 0, &target).unwrap();
    assert!(transfer.is_dependency_free());
    let transferred = transfer.commit().unwrap();
    assert_eq!(transferred.snapshot().shapes().len(), 1);
    assert_eq!(transferred.snapshot().shapes()[0].text, original.text);
    assert_eq!(
        transferred.snapshot().shapes()[0].position,
        target.text().len()
    );
    let transferred_output = transferred.snapshot().to_bytes().unwrap();
    let transferred_reopen = Document::from_bytes(&transferred_output).unwrap();
    assert_eq!(transferred_reopen.shapes()[0].geometry, original.geometry);
    let transfer_durable = transferred.patch().to_durable(durable_limits(1)).unwrap();
    let transfer_applied = target.apply_durable(&transfer_durable).unwrap();
    assert_eq!(transfer_applied.to_bytes().unwrap(), transferred_output);
    let transfer_restored = transfer_applied
        .apply_durable(&transfer_durable.inverse())
        .unwrap();
    assert!(transfer_restored.shapes().is_empty());
    assert_eq!(transfer_restored.text(), target.text());
}

#[test]
fn shape_text_edit_and_transfer_refuse_active_links() {
    let source = Document::parse(concat!(
        r#"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
        r#"{\sp{\sn hyperlink}{\sv }{\hl {\hlsrc src}{\hlloc http://example.test/x}{\hlfr Click}}}"#,
        r#"{\shptxt x}}}B}"#,
    ))
    .unwrap();
    let mut edit = source.edit();
    assert!(matches!(
        edit.set_shape_text(0, "changed"),
        Err(Error::UnsupportedSource(_))
    ));
    let target = Document::parse(r"{\rtf1 Target}").unwrap();
    assert!(matches!(
        TransferPlan::shape(&source, 0, &target),
        Err(Error::UnsupportedSource(_))
    ));
}

#[test]
fn destination_edit_refuses_unknown_syntax_without_mutating_source() {
    let source =
        Document::parse(r"{\rtf1\ansi{\*\vendor retained}\trowd\cellx1000\intbl A\cell\row}")
            .unwrap();
    let exact = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.set_table_cell_text(TableCellPath::outer(0, 0, 0), "Changed")
        .unwrap();

    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(source.to_bytes().unwrap(), exact);
    assert_eq!(source.tables()[0].rows()[0].cells()[0].text(), "A");
}

#[test]
fn transfer_is_dependency_free_and_uses_checked_ordinary_transactions() {
    let paragraph_source = Document::parse(r"{\rtf1\ansi Imported}").unwrap();
    let paragraph_target = Document::parse(r"{\rtf1\ansi Existing}").unwrap();
    let paragraph =
        TransferPlan::plain_paragraph(&paragraph_source, 0, &paragraph_target, 0).unwrap();
    assert!(paragraph.is_dependency_free());
    assert_eq!(
        paragraph.commit().unwrap().snapshot().text(),
        "Existing\nImported"
    );

    let cell_source =
        Document::parse(r"{\rtf1\ansi\trowd\cellx1000\intbl Source cell\cell\row}").unwrap();
    let cell_target =
        Document::parse(r"{\rtf1\ansi\trowd\cellx1000\intbl Target\cell\row}").unwrap();
    let transfer = TransferPlan::table_cell_text(
        &cell_source,
        &TableCellPath::outer(0, 0, 0),
        &cell_target,
        TableCellPath::outer(0, 0, 0),
    )
    .unwrap();
    assert!(transfer.is_dependency_free());
    assert_eq!(
        transfer.commit().unwrap().snapshot().tables()[0].rows()[0].cells()[0].text(),
        "Source cell"
    );
}

#[test]
fn destination_subedits_join_disjointly_and_plan_same_target_conflicts() {
    let source =
        Document::parse(r"{\rtf1\ansi\trowd\cellx1000\cellx2000\intbl A\cell B\cell\row}").unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut first = source.edit();
    first
        .set_table_cell_text(TableCellPath::outer(0, 0, 0), "First")
        .unwrap();
    let mut second = source.edit();
    second
        .set_table_cell_text(TableCellPath::outer(0, 0, 1), "Second")
        .unwrap();
    let mut joined = Composition::new(&source, limits);
    joined
        .join(first.into_sub_edit("first", limits).unwrap())
        .unwrap()
        .join(second.into_sub_edit("second", limits).unwrap())
        .unwrap();
    let committed = joined.commit().unwrap();
    assert_eq!(
        committed.snapshot().tables()[0].rows()[0].cells()[0].text(),
        "First"
    );
    assert_eq!(
        committed.snapshot().tables()[0].rows()[0].cells()[1].text(),
        "Second"
    );

    let mut left_edit = source.edit();
    left_edit
        .set_table_cell_text(TableCellPath::outer(0, 0, 0), "Left")
        .unwrap();
    let mut right_edit = source.edit();
    right_edit
        .set_table_cell_text(TableCellPath::outer(0, 0, 0), "Right")
        .unwrap();
    let mut left = Composition::new(&source, limits);
    left.join(left_edit.into_sub_edit("left-cell", limits).unwrap())
        .unwrap();
    let mut right = Composition::new(&source, limits);
    right
        .join(right_edit.into_sub_edit("right-cell", limits).unwrap())
        .unwrap();
    let plan = MergePlan::new(left, right).unwrap();
    assert_eq!(plan.conflicts().len(), 1);
    assert!(plan.finish().is_err());
    assert_eq!(source.tables()[0].rows()[0].cells()[0].text(), "A");

    let mut body_edit = source.edit();
    body_edit
        .replace_text(TextSpan::new(0, 0).unwrap(), "Body")
        .unwrap();
    let mut destination_edit = source.edit();
    destination_edit
        .set_table_cell_text(TableCellPath::outer(0, 0, 1), "Destination")
        .unwrap();
    let mut incompatible = Composition::new(&source, limits);
    incompatible
        .join(body_edit.into_sub_edit("body-domain", limits).unwrap())
        .unwrap();
    assert!(matches!(
        incompatible.join(
            destination_edit
                .into_sub_edit("destination-domain", limits)
                .unwrap()
        ),
        Err(CompositionError::Conflicts(conflicts)) if conflicts.len() == 1
    ));
}

#[test]
fn font_size_property_preserves_boundaries_defaults_and_inverse_semantics() {
    let source = Document::parse(r"{\rtf1\ansi\fs24\afs22 Alpha Beta}").unwrap();
    assert!(
        source
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );

    let mut noop = source.edit();
    noop.set_text_font_size(TextSpan::new(0, 5).unwrap(), NonZeroU16::new(24).unwrap())
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.snapshot().same_snapshot(&source));

    let mut edit = source.edit();
    edit.set_text_font_size(TextSpan::new(6, 10).unwrap(), NonZeroU16::new(23).unwrap())
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    let encoded = String::from_utf8_lossy(&bytes);
    assert!(encoded.contains(r"\fs23 "));
    assert!(encoded.contains(r"\fs24 "));
    assert!(!encoded.contains(r"\fs0"));
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), source.text());
    let size_23_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().size().get() == 23)
        .map(|run| run.text())
        .collect::<String>();
    let size_24_text = reopened
        .body()
        .runs()
        .filter(|run| run.format().size().get() == 24)
        .map(|run| run.text())
        .collect::<String>();
    assert!(size_23_text.contains("Beta"));
    assert!(size_24_text.contains("Alpha"));

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.text(), source.text());
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );
    let reopened_restored = Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened_restored.text(), source.text());

    let plain = Document::parse(r"{\rtf1\ansi Plain}").unwrap();
    assert!(
        plain
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );
    let defchp = Document::parse(r"{\rtf1\ansi{\*\defchp\fs23} Plain}").unwrap();
    assert!(
        defchp
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );

    let mut maximum = plain.edit();
    maximum
        .set_text_font_size(
            TextSpan::new(0, 5).unwrap(),
            NonZeroU16::new(65535).unwrap(),
        )
        .unwrap();
    let maximum =
        Document::from_bytes(&maximum.commit().unwrap().snapshot().to_bytes().unwrap()).unwrap();
    assert!(
        maximum
            .body()
            .runs()
            .any(|run| run.format().size().get() == 65535)
    );
}

#[test]
fn font_size_property_refuses_mixed_transport_and_malformed_durable_values() {
    use litchi_core::patch::{BlobBundle, Patch, PatchOperation, ReversibleOperation};
    use serde_json::{Number, Value};
    use std::collections::BTreeMap;

    let mixed = Document::parse(r"{\rtf1\ansi {\fs23 Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_font_size(TextSpan::new(0, 10).unwrap(), NonZeroU16::new(24).unwrap()),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed font-size state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit
            .set_text_font_size(TextSpan::new(0, 6).unwrap(), NonZeroU16::new(23).unwrap()),
        Err(Error::UnsupportedSource(
            "font-size edits require non-empty text within one paragraph"
        ))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_font_size(TextSpan::new(0, 5).unwrap(), NonZeroU16::new(23).unwrap())
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "font-size edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let mut edit = source.edit();
    edit.set_text_font_size(TextSpan::new(0, 5).unwrap(), NonZeroU16::new(23).unwrap())
        .unwrap();
    let durable = edit
        .commit()
        .unwrap()
        .patch()
        .to_durable(durable_limits(1))
        .unwrap();
    assert_eq!(durable.operations()[0].op, "character-font-size.set");
    assert_eq!(
        durable.operations()[0].preconditions["font_size_half_points"],
        Value::Number(Number::from(24_u64))
    );
    assert_eq!(
        durable.operations()[0].value,
        Value::Number(Number::from(23_u64))
    );
    let applied = source.apply_durable(&durable).unwrap();
    assert!(
        applied
            .body()
            .runs()
            .any(|run| run.format().size().get() == 23)
    );
    let restored = applied.apply_durable(&durable.inverse()).unwrap();
    assert!(
        restored
            .body()
            .runs()
            .all(|run| run.format().size().get() == 24)
    );
    Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();

    let artifact = litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex();
    let make_patch = |precondition: Value, value: Value| {
        let mut preconditions = BTreeMap::new();
        preconditions.insert(
            "artifact_sha256".to_string(),
            Value::String(artifact.clone()),
        );
        preconditions.insert("font_size_half_points".to_string(), precondition);
        let forward = PatchOperation::new(
            durable_limits(1),
            "character-font-size.set",
            "body:utf8:0-5",
            preconditions.clone(),
            value,
        )
        .unwrap();
        let inverse = PatchOperation::new(
            durable_limits(1),
            "character-font-size.set",
            "body:utf8:0-5",
            preconditions,
            Value::Number(Number::from(24_u64)),
        )
        .unwrap();
        Patch::<litchi_core::patch::Reversible>::new(
            durable_limits(1),
            "litchi-rtf",
            [ReversibleOperation::new(forward, inverse)],
            BlobBundle::new(durable_limits(1).blobs()),
            BlobBundle::new(durable_limits(1).blobs()),
        )
        .unwrap()
    };
    let malformed = [
        Value::Null,
        Value::Bool(true),
        Value::String("23".to_string()),
        Value::Number(Number::from(-1_i64)),
        Value::Number(Number::from(0_u64)),
        Value::Number(Number::from(65536_u64)),
        serde_json::json!(23.5),
    ];
    for value in malformed.iter().cloned() {
        assert!(matches!(
            source.apply_durable(&make_patch(
                Value::Number(Number::from(24_u64)),
                value.clone(),
            )),
            Err(Error::DurablePatch(_))
        ));
        assert!(matches!(
            source.apply_durable(&make_patch(value, Value::Number(Number::from(23_u64)))),
            Err(Error::DurablePatch(_))
        ));
    }
    assert!(matches!(
        source.apply_durable(&make_patch(
            Value::Number(Number::from(23_u64)),
            Value::Number(Number::from(24_u64)),
        )),
        Err(Error::StalePrecondition(
            "character font-size state differs"
        ))
    ));
}

#[test]
fn baseline_property_covers_all_states_reopen_noop_and_inverse() {
    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let states = [
        CharacterBaseline::Normal,
        CharacterBaseline::Superscript,
        CharacterBaseline::Subscript,
        CharacterBaseline::RaisedHalfPoints(8),
        CharacterBaseline::RaisedHalfPoints(31_680),
        CharacterBaseline::LoweredHalfPoints(6),
    ];

    for baseline in states {
        let mut edit = source.edit();
        edit.set_text_baseline(TextSpan::new(0, 5).unwrap(), baseline)
            .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(
            commit
                .snapshot()
                .body()
                .runs()
                .next()
                .unwrap()
                .format()
                .baseline(),
            baseline
        );
        assert_eq!(commit.snapshot().text(), source.text());

        if baseline == CharacterBaseline::Normal {
            assert!(!commit.diagnostics().changed());
            assert!(commit.snapshot().same_snapshot(&source));
        } else {
            assert!(commit.diagnostics().changed());
        }

        let reopened = Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.body().runs().next().unwrap().format().baseline(),
            baseline
        );
        let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
        assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());
    }
}

#[test]
fn baseline_property_respects_ungrouped_legacy_control_boundaries() {
    let cases = [
        (
            "super",
            CharacterBaseline::Superscript,
            CharacterBaseline::Normal,
        ),
        (
            "sub",
            CharacterBaseline::Subscript,
            CharacterBaseline::RaisedHalfPoints(8),
        ),
    ];

    for (control, initial, target) in cases {
        let source = Document::parse(&format!(r"{{\rtf1\ansi Alpha\{control} Beta}}")).unwrap();
        let mut runs = source.body().runs();
        let alpha = runs.next().unwrap();
        let beta = runs.next().unwrap();
        assert_eq!(alpha.text(), "Alpha");
        assert_eq!(alpha.format().baseline(), CharacterBaseline::Normal);
        assert_eq!(beta.text(), "Beta");
        assert_eq!(beta.format().baseline(), initial);

        let mut edit = source.edit();
        edit.set_text_baseline(TextSpan::new(5, 9).unwrap(), target)
            .unwrap();
        let commit = edit.commit().unwrap();
        let reopened = Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .body()
                .runs()
                .find(|run| run.text() == "Beta")
                .unwrap()
                .format()
                .baseline(),
            target
        );
        assert_eq!(
            commit
                .patch()
                .inverse()
                .apply(commit.snapshot())
                .unwrap()
                .to_bytes()
                .unwrap(),
            source.to_bytes().unwrap()
        );
    }
}

#[test]
fn baseline_property_preserves_unrelated_formatting_and_refuses_mixed_structure_or_opaque() {
    let source =
        Document::parse(r"{\rtf1\ansi\b\expndtw20\charscalex120\kerning8 Alpha Beta}").unwrap();
    let mut edit = source.edit();
    edit.set_text_baseline(TextSpan::new(6, 10).unwrap(), CharacterBaseline::Subscript)
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
    let alpha = reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Alpha"))
        .unwrap();
    let beta = reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Beta"))
        .unwrap();
    assert!(alpha.format().bold());
    assert_eq!(alpha.format().baseline(), CharacterBaseline::Normal);
    assert!(beta.format().bold());
    assert_eq!(beta.format().baseline(), CharacterBaseline::Subscript);
    let committed_bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        committed_bytes
            .windows(b"\\expndtw20".len())
            .any(|value| value == b"\\expndtw20")
    );
    assert!(
        committed_bytes
            .windows(b"\\charscalex120".len())
            .any(|value| value == b"\\charscalex120")
    );
    assert!(
        committed_bytes
            .windows(b"\\kerning8".len())
            .any(|value| value == b"\\kerning8")
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    let mixed = Document::parse(r"{\rtf1\ansi {\super Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_baseline(TextSpan::new(0, 10).unwrap(), CharacterBaseline::Normal),
        Err(Error::UnsupportedSource(_))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit
            .set_text_baseline(TextSpan::new(0, 6).unwrap(), CharacterBaseline::Superscript,),
        Err(Error::UnsupportedSource(_))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_baseline(TextSpan::new(0, 5).unwrap(), CharacterBaseline::Superscript)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "baseline edits refuse non-ASCII transport encodings"
        ))
    ));

    let opaque = Document::parse(r"{\rtf1\ansi Alpha{\future42 retained}}").unwrap();
    let mut opaque_edit = opaque.edit();
    opaque_edit
        .set_text_baseline(
            TextSpan::new(0, 5).unwrap(),
            CharacterBaseline::RaisedHalfPoints(8),
        )
        .unwrap();
    assert!(matches!(
        opaque_edit.commit(),
        Err(Error::UnsupportedSource(_))
    ));
}

#[test]
fn outline_property_handles_ungrouped_later_run_boundaries_and_preserves_facets() {
    let source = Document::parse(
        r"{\rtf1\ansi\b\i\uldb\strike\striked\super\scaps\caps\v\shad\embo\impr\outl Alpha \outl0 Beta}",
    )
    .unwrap();
    assert_eq!(source.text(), "Alpha Beta");

    let mut noop = source.edit();
    noop.set_text_outline(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.diagnostics().changed());
    assert!(noop_commit.snapshot().same_snapshot(&source));

    let composed_source = Document::parse(r"{\rtf1\ansi\outl Alpha \outl0 Beta}").unwrap();
    let mut composed = composed_source.edit();
    composed
        .set_text_outline(TextSpan::new(0, 6).unwrap(), true)
        .unwrap();
    composed
        .set_text_italic(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let composed_commit = composed.commit().unwrap();
    let composed_reopened =
        Document::from_bytes(&composed_commit.snapshot().to_bytes().unwrap()).unwrap();
    let alpha = composed_reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Alpha"))
        .unwrap();
    let beta = composed_reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Beta"))
        .unwrap();
    assert!(alpha.format().outline());
    assert!(!beta.format().outline());
    assert!(beta.format().italic());

    let mut edit = source.edit();
    edit.set_text_outline(TextSpan::new(6, 10).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 1);

    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\outl".len())
            .any(|window| window == br"\outl")
    );
    for control in [
        br"\shad".as_slice(),
        br"\embo".as_slice(),
        br"\impr".as_slice(),
    ] {
        assert!(bytes.windows(control.len()).any(|window| window == control));
    }

    let reopened = Document::from_bytes(&bytes).unwrap();
    let alpha = reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Alpha"))
        .unwrap()
        .format();
    let beta = reopened
        .body()
        .runs()
        .find(|run| run.text().contains("Beta"))
        .unwrap()
        .format();
    assert!(alpha.outline());
    assert!(beta.outline());
    for format in [alpha, beta] {
        assert!(format.bold());
        assert!(format.italic());
        assert_eq!(format.underline(), UnderlineStyle::Double);
        assert!(format.strike());
        assert!(format.double_strike());
        assert_eq!(format.baseline(), CharacterBaseline::Superscript);
        assert!(format.small_caps());
        assert!(format.all_caps());
        assert!(format.hidden());
    }

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.text(), source.text());
    let reopened_restored = Document::from_bytes(&restored.to_bytes().unwrap()).unwrap();
    let restored_alpha = reopened_restored
        .body()
        .runs()
        .find(|run| run.text().contains("Alpha"))
        .unwrap()
        .format();
    let restored_beta = reopened_restored
        .body()
        .runs()
        .find(|run| run.text().contains("Beta"))
        .unwrap()
        .format();
    assert!(restored_alpha.outline());
    assert!(!restored_beta.outline());
}

#[test]
fn outline_property_refuses_mixed_structure_opaque_and_non_ascii_sources() {
    let mixed = Document::parse(r"{\rtf1\ansi {\outl Alpha} Beta}").unwrap();
    let mut mixed_edit = mixed.edit();
    assert!(matches!(
        mixed_edit.set_text_outline(TextSpan::new(0, 10).unwrap(), false),
        Err(Error::UnsupportedSource(
            "the selected character span has mixed outline state"
        ))
    ));

    let paragraph = Document::parse(r"{\rtf1\ansi Alpha\par Beta}").unwrap();
    let mut paragraph_edit = paragraph.edit();
    assert!(matches!(
        paragraph_edit.set_text_outline(TextSpan::new(0, 6).unwrap(), true),
        Err(Error::UnsupportedSource(
            "outline edits require non-empty text within one paragraph"
        ))
    ));

    let opaque = Document::parse(r"{\rtf1\ansi Alpha{\future42 retained}}").unwrap();
    let mut opaque_edit = opaque.edit();
    opaque_edit
        .set_text_outline(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        opaque_edit.commit(),
        Err(Error::UnsupportedSource(_))
    ));

    let cp1252_bytes = [br"{\rtf1\ansi\ansicpg1252 Caf".as_slice(), &[0xe9], b"}"].concat();
    let cp1252 = Document::from_bytes(&cp1252_bytes).unwrap();
    let mut cp1252_edit = cp1252.edit();
    cp1252_edit
        .set_text_outline(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        cp1252_edit.commit(),
        Err(Error::UnsupportedSource(
            "outline edits refuse non-ASCII transport encodings"
        ))
    ));

    let source = Document::parse(r"{\rtf1\ansi Alpha Beta}").unwrap();
    let mut conflict = source.edit();
    conflict
        .set_text_outline(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    assert!(matches!(
        conflict.set_text_outline(TextSpan::new(4, 10).unwrap(), true),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
}

#[test]
fn outline_property_has_durable_schema_replay_stale_and_malformed_value_guards() {
    use litchi_core::patch::Patch;
    use serde_json::Value;

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let mut edit = source.edit();
    edit.set_text_outline(TextSpan::new(0, 5).unwrap(), true)
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let operation = &durable.operations()[0];
    assert_eq!(operation.op, "character-outline.set");
    assert_eq!(operation.target, "body:utf8:0-5");
    assert_eq!(operation.preconditions["outline"], Value::Bool(false));
    assert_eq!(operation.value, Value::Bool(true));

    let encoded = durable.to_deterministic_json().unwrap();
    assert_eq!(encoded, durable.to_deterministic_json().unwrap());
    let decoded = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &encoded,
        durable_limits(1),
    )
    .unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert!(applied.body().runs().all(|run| run.format().outline()));
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert!(restored.body().runs().all(|run| !run.format().outline()));

    let stale_source = {
        let mut stale_edit = source.edit();
        stale_edit
            .set_text_outline(TextSpan::new(0, 5).unwrap(), true)
            .unwrap();
        stale_edit.commit().unwrap().into_snapshot()
    };
    let stale_bytes = stale_source.to_bytes().unwrap();
    let mut stale_json: Value = serde_json::from_slice(&encoded).unwrap();
    stale_json["operations"][0]["forward"]["preconditions"]["artifact_sha256"] =
        Value::String(litchi_core::patch::BlobId::of(&stale_bytes).as_hex());
    let stale_encoded = serde_json::to_vec(&stale_json).unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &stale_encoded,
        durable_limits(1),
    )
    .unwrap();
    assert!(matches!(
        stale_source.apply_durable(&stale),
        Err(Error::StalePrecondition(_))
    ));

    for malformed in [Value::Null, Value::String("true".to_string())] {
        let mut malformed_json: Value = serde_json::from_slice(&encoded).unwrap();
        malformed_json["operations"][0]["forward"]["value"] = malformed;
        let malformed_json = serde_json::to_vec(&malformed_json).unwrap();
        let malformed = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &malformed_json,
            durable_limits(1),
        )
        .unwrap();
        assert!(matches!(
            source.apply_durable(&malformed),
            Err(Error::DurablePatch(_))
        ));
    }

    let mut malformed_precondition: Value = serde_json::from_slice(&encoded).unwrap();
    malformed_precondition["operations"][0]["forward"]["preconditions"]["outline"] = Value::Null;
    let malformed_precondition = serde_json::to_vec(&malformed_precondition).unwrap();
    let malformed_precondition = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &malformed_precondition,
        durable_limits(1),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&malformed_precondition),
        Err(Error::DurablePatch(_))
    ));
}

#[test]
fn baseline_property_has_durable_schema_replay_stale_and_malformed_value_guards() {
    use litchi_core::patch::Patch;
    use serde_json::Value;

    let source = Document::parse(r"{\rtf1\ansi Alpha}").unwrap();
    let mut edit = source.edit();
    edit.set_text_baseline(
        TextSpan::new(0, 5).unwrap(),
        CharacterBaseline::RaisedHalfPoints(8),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().to_durable(durable_limits(1)).unwrap();
    let operation = &durable.operations()[0];
    assert_eq!(operation.op, "character-baseline.set");
    assert_eq!(operation.target, "body:utf8:0-5");
    assert!(operation.preconditions.contains_key("baseline"));
    assert_ne!(operation.value, Value::Null);

    let encoded = durable.to_deterministic_json().unwrap();
    let decoded = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &encoded,
        durable_limits(1),
    )
    .unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert_eq!(
        applied.body().runs().next().unwrap().format().baseline(),
        CharacterBaseline::RaisedHalfPoints(8)
    );
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(
        restored.body().runs().next().unwrap().format().baseline(),
        CharacterBaseline::Normal
    );

    let stale_source = {
        let mut stale_edit = source.edit();
        stale_edit
            .set_text_baseline(TextSpan::new(0, 5).unwrap(), CharacterBaseline::Subscript)
            .unwrap();
        stale_edit.commit().unwrap().into_snapshot()
    };
    let stale_bytes = stale_source.to_bytes().unwrap();
    let mut stale_json: Value = serde_json::from_slice(&encoded).unwrap();
    stale_json["operations"][0]["forward"]["preconditions"]["artifact_sha256"] =
        Value::String(litchi_core::patch::BlobId::of(&stale_bytes).as_hex());
    let stale_encoded = serde_json::to_vec(&stale_json).unwrap();
    let stale = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &stale_encoded,
        durable_limits(1),
    )
    .unwrap();
    assert!(matches!(
        stale_source.apply_durable(&stale),
        Err(Error::StalePrecondition(_))
    ));

    let mut malformed_value: Value = serde_json::from_slice(&encoded).unwrap();
    malformed_value["operations"][0]["forward"]["value"] = Value::Null;
    let malformed_value = serde_json::to_vec(&malformed_value).unwrap();
    let malformed_value = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &malformed_value,
        durable_limits(1),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&malformed_value),
        Err(Error::DurablePatch(_))
    ));

    for malformed in [
        "unknown",
        "raised-half-points:0",
        "raised-half-points:01",
        "raised-half-points:31681",
        "lowered-half-points:0",
    ] {
        let mut malformed_value: Value = serde_json::from_slice(&encoded).unwrap();
        malformed_value["operations"][0]["forward"]["value"] = Value::String(malformed.to_string());
        let malformed_value = serde_json::to_vec(&malformed_value).unwrap();
        let malformed_value = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &malformed_value,
            durable_limits(1),
        )
        .unwrap();
        assert!(matches!(
            source.apply_durable(&malformed_value),
            Err(Error::DurablePatch(_))
        ));
    }

    let mut malformed_precondition: Value = serde_json::from_slice(&encoded).unwrap();
    malformed_precondition["operations"][0]["forward"]["preconditions"]["baseline"] = Value::Null;
    let malformed_precondition = serde_json::to_vec(&malformed_precondition).unwrap();
    let malformed_precondition = Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
        &malformed_precondition,
        durable_limits(1),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&malformed_precondition),
        Err(Error::DurablePatch(_))
    ));
}

#[test]
fn character_property_edits_preserve_paragraph_formatting() {
    let source = Document::parse(r"{\rtf1\ansi\sb120\pagebb Alpha}").unwrap();
    let exact = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.set_text_baseline(TextSpan::new(0, 5).unwrap(), CharacterBaseline::Superscript)
        .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.snapshot().to_bytes().unwrap();
    assert!(
        bytes
            .windows(br"\sb120".len())
            .any(|value| value == br"\sb120")
    );
    assert!(
        bytes
            .windows(br"\pagebb".len())
            .any(|value| value == br"\pagebb")
    );
    let reopened = Document::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.body().runs().next().unwrap().format().baseline(),
        CharacterBaseline::Superscript
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .to_bytes()
            .unwrap(),
        exact
    );
}
