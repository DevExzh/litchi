#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Alignment, CharacterBaseline, Document, ProtectionType,
    edit::{
        Composition, CompositionError, CompositionLimits, Error, Limits, ParagraphLayoutPatch,
        ParagraphLayoutUpdate, TextSpan,
    },
    transport::compress,
};

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(1, 1024 * 1024, 1024 * 1024),
        1024 * 1024,
        max_operations,
        8,
        256 * 1024,
        512 * 1024,
    )
}

fn layout(document: &Document, position: usize) -> litchi_rtf::edit::ParagraphLayout {
    document
        .body()
        .paragraphs()
        .nth(position)
        .unwrap()
        .format()
        .layout()
}

#[test]
fn first_middle_last_batch_preserves_text_runs_and_unrelated_paragraph_properties() {
    let source = Document::parse(concat!(
        r"{\rtf1\ansi ",
        r"\pard\qc\sb120\sa40\li240\ri60\fi-30\keep \b First\b0\par ",
        r"\pard\qr\sb80\sa20\li120\sl300\slmult1 \i Second\i0\par ",
        r"\pard\qj\sb60\sa10\ri90\pagebb Third}"
    ))
    .unwrap();
    let before_text = source.text().to_string();
    let before_runs = source
        .body()
        .runs()
        .map(|run| {
            (
                run.text().to_string(),
                run.format().bold(),
                run.format().italic(),
                run.format().underline(),
            )
        })
        .collect::<Vec<_>>();
    let before_formats = source
        .body()
        .paragraphs()
        .map(|paragraph| {
            (
                paragraph.format().alignment(),
                paragraph.format().spacing().line,
                paragraph.format().spacing().line_multiple,
            )
        })
        .collect::<Vec<_>>();

    let updates = [
        ParagraphLayoutUpdate::new(
            0,
            ParagraphLayoutPatch::new()
                .clear_space_before()
                .clear_keep_together(),
        ),
        ParagraphLayoutUpdate::new(
            1,
            ParagraphLayoutPatch::new()
                .with_left_indent(360)
                .with_keep_with_next(true),
        ),
        ParagraphLayoutUpdate::new(
            2,
            ParagraphLayoutPatch::new()
                .clear_right_indent()
                .with_first_line_indent(120)
                .clear_page_break_before(),
        ),
    ];
    let mut edit = source.edit();
    edit.patch_body_paragraph_layouts(&updates).unwrap();
    let commit = edit.commit().unwrap();
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

    assert_eq!(commit.snapshot().text(), before_text);
    assert_eq!(layout(commit.snapshot(), 0).space_before(), 0);
    assert!(!layout(commit.snapshot(), 0).keep_together());
    assert_eq!(layout(commit.snapshot(), 1).left_indent(), 360);
    assert!(layout(commit.snapshot(), 1).keep_with_next());
    assert_eq!(layout(commit.snapshot(), 2).right_indent(), 0);
    assert_eq!(layout(commit.snapshot(), 2).first_line_indent(), 120);
    assert!(!layout(commit.snapshot(), 2).page_break_before());
    let after_runs = commit
        .snapshot()
        .body()
        .runs()
        .map(|run| {
            (
                run.text().to_string(),
                run.format().bold(),
                run.format().italic(),
                run.format().underline(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(after_runs, before_runs);
    let after_formats = commit
        .snapshot()
        .body()
        .paragraphs()
        .map(|paragraph| {
            (
                paragraph.format().alignment(),
                paragraph.format().spacing().line,
                paragraph.format().spacing().line_multiple,
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(after_formats, before_formats);
    let reopened = Document::from_bytes(&commit.snapshot().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), before_text);
    assert_eq!(layout(&reopened, 1).left_indent(), 360);
}

#[test]
fn batch_preflights_shape_selectors_limits_and_conflicts_atomically() {
    let source = Document::parse(r"{\rtf1\ansi One\par Two\par Three}").unwrap();

    let mut empty = source.edit();
    assert!(matches!(
        empty.patch_body_paragraph_layouts(&[]),
        Err(Error::EmptyParagraphLayoutBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    let mut empty_patch = source.edit();
    assert!(matches!(
        empty_patch.patch_paragraph_layout(1, ParagraphLayoutPatch::new()),
        Err(Error::EmptyParagraphLayoutPatch { position: 1 })
    ));
    assert_eq!(empty_patch.operation_count(), 0);

    for updates in [
        [
            ParagraphLayoutUpdate::new(2, ParagraphLayoutPatch::new().with_space_before(20)),
            ParagraphLayoutUpdate::new(1, ParagraphLayoutPatch::new().with_space_before(10)),
        ],
        [
            ParagraphLayoutUpdate::new(1, ParagraphLayoutPatch::new().with_space_before(20)),
            ParagraphLayoutUpdate::new(1, ParagraphLayoutPatch::new().with_space_after(10)),
        ],
    ] {
        let mut edit = source.edit();
        assert!(matches!(
            edit.patch_body_paragraph_layouts(&updates),
            Err(Error::ParagraphLayoutBatchOutOfOrder { .. })
        ));
        assert_eq!(edit.operation_count(), 0);
    }

    let mut late = source.edit();
    assert!(matches!(
        late.patch_body_paragraph_layouts(&[
            ParagraphLayoutUpdate::new(0, ParagraphLayoutPatch::new().with_space_before(20),),
            ParagraphLayoutUpdate::new(9, ParagraphLayoutPatch::new().with_space_after(10),),
        ]),
        Err(Error::ParagraphOutOfRange {
            position: 9,
            count: 3
        })
    ));
    assert_eq!(late.operation_count(), 0);

    let mut bounded = source.edit_with_limits(Limits::new(1));
    assert!(matches!(
        bounded.patch_body_paragraph_layouts(&[
            ParagraphLayoutUpdate::new(0, ParagraphLayoutPatch::new().with_space_before(20),),
            ParagraphLayoutUpdate::new(2, ParagraphLayoutPatch::new().with_space_after(10),),
        ]),
        Err(Error::OperationLimit {
            observed: 2,
            limit: 1
        })
    ));
    assert_eq!(bounded.operation_count(), 0);

    let mut overlap = source.edit();
    overlap
        .patch_paragraph_layout(1, ParagraphLayoutPatch::new().with_space_before(20))
        .unwrap();
    assert!(matches!(
        overlap.patch_paragraph_layout(1, ParagraphLayoutPatch::new().with_space_before(40)),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    overlap
        .patch_paragraph_layout(1, ParagraphLayoutPatch::new().with_space_after(30))
        .unwrap();
    assert_eq!(overlap.operation_count(), 2);
}

#[test]
fn layout_composes_with_alignment_and_disjoint_prepared_work_but_not_text() {
    let source = Document::parse(r"{\rtf1\ansi One\par Two}").unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_keep_together(true))
        .unwrap();
    edit.set_paragraph_alignment(0, Alignment::Center).unwrap();
    assert!(matches!(
        edit.replace_text(TextSpan::new(0, 3).unwrap(), "First"),
        Err(Error::ParagraphLayoutTextConflict)
    ));
    let commit = edit.commit().unwrap();
    assert!(layout(commit.snapshot(), 0).keep_together());
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

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut before = source.edit();
    before
        .patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(120))
        .unwrap();
    let mut after = source.edit();
    after
        .patch_paragraph_layout(1, ParagraphLayoutPatch::new().with_space_after(80))
        .unwrap();
    let mut composition = Composition::new(&source, limits);
    composition
        .join(before.into_sub_edit("before", limits).unwrap())
        .unwrap()
        .join(after.into_sub_edit("after", limits).unwrap())
        .unwrap();
    let composed = composition.commit().unwrap();
    assert_eq!(layout(composed.snapshot(), 0).space_before(), 120);
    assert_eq!(layout(composed.snapshot(), 1).space_after(), 80);

    let mut left = source.edit();
    left.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_left_indent(100))
        .unwrap();
    let mut right = source.edit();
    right
        .patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_left_indent(200))
        .unwrap();
    let mut conflict = Composition::new(&source, limits);
    conflict
        .join(left.into_sub_edit("left", limits).unwrap())
        .unwrap();
    assert!(matches!(
        conflict.join(right.into_sub_edit("right", limits).unwrap()),
        Err(CompositionError::Conflicts(conflicts)) if !conflicts.is_empty()
    ));

    let mut paragraph_layout = source.edit();
    paragraph_layout
        .patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(120))
        .unwrap();
    let mut character = source.edit();
    character
        .set_text_baseline(TextSpan::new(0, 3).unwrap(), CharacterBaseline::Superscript)
        .unwrap();
    let mut cross_kind_conflict = Composition::new(&source, limits);
    cross_kind_conflict
        .join(
            paragraph_layout
                .into_sub_edit("paragraph-layout", limits)
                .unwrap(),
        )
        .unwrap();
    assert!(matches!(
        cross_kind_conflict.join(character.into_sub_edit("character", limits).unwrap()),
        Err(CompositionError::Conflicts(conflicts)) if !conflicts.is_empty()
    ));
}

#[test]
fn durable_layout_patch_is_deterministic_reversible_and_stale_checked() {
    let source = Document::parse(r"{\rtf1\ansi\sb40 One\par\pard\li80 Two}").unwrap();
    let mut edit = source.edit();
    edit.patch_body_paragraph_layouts(&[
        ParagraphLayoutUpdate::new(
            0,
            ParagraphLayoutPatch::new()
                .with_space_before(120)
                .with_keep_together(true),
        ),
        ParagraphLayoutUpdate::new(
            1,
            ParagraphLayoutPatch::new()
                .with_left_indent(240)
                .with_page_break_before(true),
        ),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored.to_bytes().unwrap(),
        source.to_bytes().unwrap(),
        "source-bound inverse must restore the exact source artifact"
    );
    assert_eq!(restored.text(), source.text());
    assert_eq!(layout(&restored, 0), layout(&source, 0));
    assert_eq!(layout(&restored, 1), layout(&source, 1));

    assert!(matches!(
        commit.patch().to_durable(durable_limits(2)),
        Err(litchi_core::patch::PatchError::InvalidText {
            field: "RTF paragraph-layout durable patches are not supported"
        })
    ));

    use litchi_core::patch::{BlobBundle, Patch as CorePatch, PatchOperation, ReversibleOperation};
    use serde_json::json;
    use std::collections::BTreeMap;
    let limits = durable_limits(1);
    let mut preconditions = BTreeMap::new();
    preconditions.insert(
        "artifact_sha256".to_string(),
        serde_json::Value::String(
            litchi_core::patch::BlobId::of(&source.to_bytes().unwrap()).as_hex(),
        ),
    );
    preconditions.insert("layout".to_string(), json!({"space_before": 40}));
    let forward = PatchOperation::new(
        limits,
        "paragraph-layout.patch",
        "body:paragraph:0",
        preconditions.clone(),
        json!({"space_before": 120}),
    )
    .unwrap();
    let inverse = PatchOperation::new(
        limits,
        "paragraph-layout.patch",
        "body:paragraph:0",
        preconditions,
        json!({"space_before": 40}),
    )
    .unwrap();
    let unsupported = CorePatch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [ReversibleOperation::new(forward, inverse)],
        BlobBundle::new(limits.blobs()),
        BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&unsupported),
        Err(Error::DurablePatch(message))
            if message == "paragraph-layout durable patches are not supported"
    ));
}

#[test]
fn exact_noops_and_source_closure_refusals_do_not_publish() {
    let protected = Document::parse(r"{\rtf1\ansi\readprot\enforceprot1\sb40 Same}").unwrap();
    let exact = protected.to_bytes().unwrap();
    let mut noop = protected.edit();
    noop.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(40))
        .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.snapshot().to_bytes().unwrap(), exact);

    let mut changed = protected.edit();
    changed
        .patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(80))
        .unwrap();
    assert!(matches!(
        changed.commit(),
        Err(Error::ProtectedDocument {
            protection_type: ProtectionType::ReadOnly
        })
    ));

    for source in [
        Document::parse(r"{\rtf1\ansi A\future42 B}").unwrap(),
        Document::parse(r"{\rtf1\ansi{\stylesheet{\s1 Named;}}\s1 Styled}").unwrap(),
        Document::parse(r"{\rtf1\ansi\ls1 Listed}").unwrap(),
        Document::parse(r"{\rtf1\ansi One\line Two}").unwrap(),
        Document::parse(r"{\rtf1\ansi\trowd\cellx1000\intbl Cell\cell\row}").unwrap(),
    ] {
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        assert!(matches!(
            edit.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(20)),
            Err(Error::UnsupportedSource(_))
        ));
        assert_eq!(edit.operation_count(), 0);
        assert_eq!(source.to_bytes().unwrap(), exact);
    }

    let raw = br"{\rtf1\ansi One}";
    let compressed = Document::from_bytes(&compress(raw, true).unwrap()).unwrap();
    let mut edit = compressed.edit();
    assert!(matches!(
        edit.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(20)),
        Err(Error::UnsupportedSource(_))
    ));

    let mut non_ascii_bytes = br"{\rtf1\ansi\ansicpg1252 Caf".to_vec();
    non_ascii_bytes.push(0xe9);
    non_ascii_bytes.push(b'}');
    let non_ascii = Document::from_bytes(&non_ascii_bytes).unwrap();
    let mut edit = non_ascii.edit();
    assert!(matches!(
        edit.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(20)),
        Err(Error::UnsupportedSource(_))
    ));

    let source_text = r"{\rtf1 A}";
    let limits = litchi_rtf::read::Limits::new().with_max_source_bytes(source_text.len());
    let bounded = Document::parse_with_limits(source_text, limits).unwrap();
    let exact = bounded.to_bytes().unwrap();
    let mut edit = bounded.edit();
    edit.patch_paragraph_layout(0, ParagraphLayoutPatch::new().with_space_before(20))
        .unwrap();
    assert!(matches!(edit.commit(), Err(Error::Write(_))));
    assert_eq!(bounded.to_bytes().unwrap(), exact);
}
