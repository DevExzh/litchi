#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Alignment, Document, HeaderFooterType, TableCellPath,
    edit::{
        Composition, CompositionError, CompositionLimits, Error, HeaderFooterParagraph, History,
        HistoryLimits, Limits, MergePlan, MergeResolution, TextSpan, TransferPlan,
    },
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
