#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::Position;
use litchi_rtf::raw::Tree;
use litchi_rtf::{
    Alignment, CharacterBaseline, EditOptions, MAX_PARAGRAPH_FRAME_TWIPS, OpenOptions, Package,
    ParagraphFrame, ParagraphFrameHorizontalPosition, ParagraphFrameHorizontalReference,
    ParagraphFrameTextFlow, ParagraphFrameVerticalPosition, ParagraphFrameVerticalReference,
    ParagraphFrameWrap, SemanticLimitKind, SemanticLimits, SourceLimits,
    edit::{
        Composition, CompositionError, CompositionLimits, Error, ParagraphLayoutPatch,
        ParagraphLayoutUpdate, TextSpan,
    },
    transport::{CompressionLimits, compress},
};

fn semantic_limits(max_operations: usize) -> SemanticLimits {
    SemanticLimits::builder()
        .max_operations(max_operations)
        .unwrap()
        .build()
        .unwrap()
}

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

fn layout(document: &Package, position: usize) -> litchi_rtf::edit::ParagraphLayout {
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
    let source = Package::parse(concat!(
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
            Position::new(0),
            ParagraphLayoutPatch::new()
                .clear_space_before()
                .clear_keep_together(),
        ),
        ParagraphLayoutUpdate::new(
            Position::new(1),
            ParagraphLayoutPatch::new()
                .with_left_indent(360)
                .with_keep_with_next(true),
        ),
        ParagraphLayoutUpdate::new(
            Position::new(2),
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
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    assert_eq!(commit.package().text(), before_text);
    assert_eq!(layout(commit.package(), 0).space_before(), 0);
    assert!(!layout(commit.package(), 0).keep_together());
    assert_eq!(layout(commit.package(), 1).left_indent(), 360);
    assert!(layout(commit.package(), 1).keep_with_next());
    assert_eq!(layout(commit.package(), 2).right_indent(), 0);
    assert_eq!(layout(commit.package(), 2).first_line_indent(), 120);
    assert!(!layout(commit.package(), 2).page_break_before());
    let after_runs = commit
        .package()
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
        .package()
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
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), before_text);
    assert_eq!(layout(&reopened, 1).left_indent(), 360);
}

#[test]
fn batch_preflights_shape_selectors_limits_and_conflicts_atomically() {
    let source = Package::parse(r"{\rtf1\ansi One\par Two\par Three}").unwrap();

    let mut empty = source.edit();
    assert!(matches!(
        empty.patch_body_paragraph_layouts(&[]),
        Err(Error::EmptyParagraphLayoutBatch)
    ));
    assert_eq!(empty.operation_count(), 0);

    let mut empty_patch = source.edit();
    assert!(matches!(
        empty_patch.patch_paragraph_layout(Position::new(1), ParagraphLayoutPatch::new()),
        Err(Error::EmptyParagraphLayoutPatch { position: 1 })
    ));
    assert_eq!(empty_patch.operation_count(), 0);

    for updates in [
        [
            ParagraphLayoutUpdate::new(
                Position::new(2),
                ParagraphLayoutPatch::new().with_space_before(20),
            ),
            ParagraphLayoutUpdate::new(
                Position::new(1),
                ParagraphLayoutPatch::new().with_space_before(10),
            ),
        ],
        [
            ParagraphLayoutUpdate::new(
                Position::new(1),
                ParagraphLayoutPatch::new().with_space_before(20),
            ),
            ParagraphLayoutUpdate::new(
                Position::new(1),
                ParagraphLayoutPatch::new().with_space_after(10),
            ),
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
            ParagraphLayoutUpdate::new(
                Position::new(0),
                ParagraphLayoutPatch::new().with_space_before(20),
            ),
            ParagraphLayoutUpdate::new(
                Position::new(9),
                ParagraphLayoutPatch::new().with_space_after(10),
            ),
        ]),
        Err(Error::ParagraphOutOfRange {
            position: 9,
            count: 3
        })
    ));
    assert_eq!(late.operation_count(), 0);

    let mut bounded =
        source.edit_with_options(EditOptions::new().semantic_limits(semantic_limits(1)));
    assert!(matches!(
        bounded.patch_body_paragraph_layouts(&[
            ParagraphLayoutUpdate::new(
                Position::new(0),
                ParagraphLayoutPatch::new().with_space_before(20),
            ),
            ParagraphLayoutUpdate::new(
                Position::new(2),
                ParagraphLayoutPatch::new().with_space_after(10),
            ),
        ]),
        Err(Error::Limit {
            kind: SemanticLimitKind::Operations,
            observed: 2,
            maximum: 1
        })
    ));
    assert_eq!(bounded.operation_count(), 0);

    let mut overlap = source.edit();
    overlap
        .patch_paragraph_layout(
            Position::new(1),
            ParagraphLayoutPatch::new().with_space_before(20),
        )
        .unwrap();
    assert!(matches!(
        overlap.patch_paragraph_layout(
            Position::new(1),
            ParagraphLayoutPatch::new().with_space_before(40)
        ),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    overlap
        .patch_paragraph_layout(
            Position::new(1),
            ParagraphLayoutPatch::new().with_space_after(30),
        )
        .unwrap();
    assert_eq!(overlap.operation_count(), 2);
}

#[test]
fn layout_composes_with_alignment_and_disjoint_prepared_work_but_not_text() {
    let source = Package::parse(r"{\rtf1\ansi One\par Two}").unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_keep_together(true),
    )
    .unwrap();
    edit.set_paragraph_alignment(Position::new(0), Alignment::Center)
        .unwrap();
    assert!(matches!(
        edit.replace_text(TextSpan::new(0, 3).unwrap(), "First"),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation && refusal.reason() == "paragraph-layout updates cannot compose with body text or character changes"
    ));
    let commit = edit.commit().unwrap();
    assert!(layout(commit.package(), 0).keep_together());
    assert_eq!(
        commit
            .package()
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
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(120),
        )
        .unwrap();
    let mut after = source.edit();
    after
        .patch_paragraph_layout(
            Position::new(1),
            ParagraphLayoutPatch::new().with_space_after(80),
        )
        .unwrap();
    let mut composition = Composition::new(&source, limits);
    composition
        .join(before.into_sub_edit("before", limits).unwrap())
        .unwrap()
        .join(after.into_sub_edit("after", limits).unwrap())
        .unwrap();
    let composed = composition.commit().unwrap();
    assert_eq!(layout(composed.package(), 0).space_before(), 120);
    assert_eq!(layout(composed.package(), 1).space_after(), 80);

    let mut left = source.edit();
    left.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_left_indent(100),
    )
    .unwrap();
    let mut right = source.edit();
    right
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_left_indent(200),
        )
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
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(120),
        )
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
    let source = Package::parse(r"{\rtf1\ansi\sb40 One\par\pard\li80 Two}").unwrap();
    let mut edit = source.edit();
    edit.patch_body_paragraph_layouts(&[
        ParagraphLayoutUpdate::new(
            Position::new(0),
            ParagraphLayoutPatch::new()
                .with_space_before(120)
                .with_keep_together(true),
        ),
        ParagraphLayoutUpdate::new(
            Position::new(1),
            ParagraphLayoutPatch::new()
                .with_left_indent(240)
                .with_page_break_before(true),
        ),
    ])
    .unwrap();
    let commit = edit.commit().unwrap();
    let restored = commit
        .package()
        .apply(&commit.patch().inverse())
        .unwrap()
        .into_package();
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
fn source_bound_frame_layout_crud_is_reversible_and_source_checked() {
    let source = Package::parse(r"{\rtf1\ansi One\par Two}").unwrap();
    let frame = ParagraphFrame {
        width_twips: Some(720),
        height_twips: Some(-480),
        horizontal_reference: ParagraphFrameHorizontalReference::Margin,
        horizontal_position: ParagraphFrameHorizontalPosition::NegativeOffset(-80),
        vertical_reference: ParagraphFrameVerticalReference::Page,
        vertical_position: ParagraphFrameVerticalPosition::Offset(120),
        anchor_locked: Some(true),
        no_wrap: true,
        horizontal_text_distance_twips: Some(40),
        horizontal_text_offset_twips: Some(10),
        vertical_text_offset_twips: Some(20),
        wrap: ParagraphFrameWrap::Around,
        overlay: true,
        no_overlap: Some(true),
        text_flow: ParagraphFrameTextFlow::TopToBottomRightToLeft,
    };
    let mut create = source.edit();
    create
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_frame(frame),
        )
        .unwrap();
    let created = create.commit().unwrap();
    assert_eq!(layout(created.package(), 0).frame(), Some(frame));
    let created_bytes = created.package().to_bytes().unwrap();
    let created_text = String::from_utf8(created_bytes.clone()).unwrap();
    assert!(created_text.contains(r"\absw720"));
    assert!(created_text.contains(r"\posnegx-80"));
    let reopened = Package::from_bytes(&created_bytes).unwrap();
    assert_eq!(layout(&reopened, 0).frame(), Some(frame));
    assert_eq!(reopened.text(), source.text());
    assert_eq!(
        created
            .package()
            .apply(&created.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    let updated_frame = ParagraphFrame {
        width_twips: Some(960),
        horizontal_position: ParagraphFrameHorizontalPosition::Right,
        wrap: ParagraphFrameWrap::Through,
        ..frame
    };
    let mut update = created.package().edit();
    update
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_frame(updated_frame),
        )
        .unwrap();
    let updated = update.commit().unwrap();
    assert_eq!(layout(updated.package(), 0).frame(), Some(updated_frame));
    assert!(matches!(
        source.apply(updated.patch()),
        Err(Error::PatchConflict)
    ));

    let mut remove = updated.package().edit();
    remove
        .patch_paragraph_layout(Position::new(0), ParagraphLayoutPatch::new().clear_frame())
        .unwrap();
    let removed = remove.commit().unwrap();
    assert_eq!(layout(removed.package(), 0).frame(), None);
    assert_eq!(removed.package().text(), source.text());
    assert_eq!(
        removed
            .package()
            .apply(&removed.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        updated.package().to_bytes().unwrap()
    );
}

#[test]
fn frame_layout_batch_validation_and_unknown_source_spans_are_atomic() {
    let source = Package::parse(r"{\rtf1\ansi One\par Two}").unwrap();
    let before = source.to_bytes().unwrap();
    let invalid = ParagraphFrame {
        width_twips: Some(MAX_PARAGRAPH_FRAME_TWIPS as u32 + 1),
        ..ParagraphFrame::default()
    };
    let mut batch = source.edit();
    assert!(matches!(
        batch.patch_body_paragraph_layouts(&[
            ParagraphLayoutUpdate::new(
                Position::new(0),
                ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
            ),
            ParagraphLayoutUpdate::new(
                Position::new(1),
                ParagraphLayoutPatch::new().with_frame(invalid)
            ),
        ]),
        Err(Error::Rtf(_))
    ));
    assert_eq!(batch.operation_count(), 0);
    assert_eq!(source.to_bytes().unwrap(), before);

    let unknown = Package::parse(r"{\rtf1\ansi\future42 One{\*\vendor retained}\par Two}").unwrap();
    let unknown_before = unknown.to_bytes().unwrap();
    let mut unknown_edit = unknown.edit();
    unknown_edit
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
        )
        .unwrap();
    let unknown_commit = unknown_edit.commit().unwrap();
    let unknown_after = unknown_commit.package().to_bytes().unwrap();
    assert!(
        unknown_after
            .windows(br"\future42".len())
            .any(|window| { window == br"\future42" })
    );
    assert_eq!(
        unknown_after
            .windows(br"\future42".len())
            .filter(|window| *window == br"\future42")
            .count(),
        1
    );
    assert!(
        unknown_after
            .windows(br"{\*\vendor retained}".len())
            .any(|window| window == br"{\*\vendor retained}")
    );
    assert_eq!(unknown_commit.package().opaque(), unknown.opaque());
    assert_eq!(unknown_commit.package().text(), unknown.text());
    assert_eq!(
        unknown_commit
            .package()
            .apply(&unknown_commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        unknown_before
    );

    let metadata =
        Package::parse(r"{\rtf1\ansi{\info{\*\futuremeta vendor-byte-payload}} One}").unwrap();
    let metadata_before = metadata.to_bytes().unwrap();
    let mut metadata_edit = metadata.edit();
    metadata_edit
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(20),
        )
        .unwrap();
    let metadata_after = metadata_edit
        .commit()
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    let metadata_prefix = br"{\rtf1\ansi{\info{\*\futuremeta vendor-byte-payload}}";
    assert!(metadata_after.starts_with(metadata_prefix));
    assert!(
        metadata_after
            .windows(br"\sb20".len())
            .any(|window| window == br"\sb20")
    );
    assert_eq!(
        &metadata_before[..metadata_prefix.len()],
        &metadata_after[..metadata_prefix.len()]
    );
}

#[test]
fn frame_layout_preserves_zero_width_bookmark_events() {
    let source = Package::parse(r#"{\rtf1\ansi One{\*\bkmkstart bm}two{\*\bkmkend bm}}"#).unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.package().to_bytes().unwrap();
    assert!(String::from_utf8_lossy(&bytes).contains(r"\bkmkstart bm"));
    let reopened = Tree::parse_bytes(&bytes, OpenOptions::default()).unwrap();
    assert_eq!(reopened.bookmarks().bookmarks().len(), 1);
    assert_eq!(reopened.bookmarks().bookmarks()[0].name, "bm");
    assert_eq!(reopened.bookmarks().bookmarks()[0].position, 3);
    assert_eq!(reopened.bookmarks().bookmarks()[0].content, "two");
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn frame_layout_does_not_duplicate_leading_zero_width_bookmark_events() {
    let source =
        Package::parse(r#"{\rtf1\ansi{\*\bkmkstart leading}One{\*\bkmkend leading}}"#).unwrap();
    let before = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.package().to_bytes().unwrap();
    assert_eq!(
        bytes
            .windows(br"\bkmkstart leading".len())
            .filter(|window| *window == br"\bkmkstart leading")
            .count(),
        1
    );
    let reopened = Tree::parse_bytes(&bytes, OpenOptions::default()).unwrap();
    assert_eq!(reopened.bookmarks().bookmarks().len(), 1);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn frame_layout_preserves_offset_only_break_events() {
    let source = Package::parse(r#"{\rtf1\ansi One\page Two\column Three\softline Four}"#).unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.package().to_bytes().unwrap();
    let serialized = String::from_utf8_lossy(&bytes);
    assert!(serialized.contains(r"\page "));
    assert!(serialized.contains(r"\column "));
    assert!(serialized.contains(r"\softline "));
    let reopened = Tree::parse_bytes(&bytes, OpenOptions::default()).unwrap();
    assert_eq!(reopened.body_story_events().len(), 3);
    assert_eq!(reopened.text(), "OneTwoThreeFour");
}

#[test]
fn frame_layout_preserves_leading_body_opaque_destination() {
    let source = Package::parse(r#"{\rtf1\ansi{\*\vendor leading}One}"#).unwrap();
    let before = source.to_bytes().unwrap();
    assert_eq!(source.opaque().len(), 1);
    assert!(matches!(
        source.opaque()[0].anchor(),
        litchi_rtf::opaque::Anchor::Body(0)
    ));
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_space_before(20),
    )
    .unwrap();
    let committed = edit.commit().unwrap();
    let after = committed.package().to_bytes().unwrap();
    assert!(
        after
            .windows(br"{\*\vendor leading}".len())
            .any(|window| { window == br"{\*\vendor leading}" })
    );
    assert_eq!(committed.package().opaque(), source.opaque());
    assert_eq!(
        committed
            .package()
            .apply(&committed.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        before
    );
}

#[test]
fn exact_noops_and_source_closure_refusals_do_not_publish() {
    let protected = Package::parse(r"{\rtf1\ansi\readprot\enforceprot1\sb40 Same}").unwrap();
    let exact = protected.to_bytes().unwrap();
    let mut noop = protected.edit();
    noop.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_space_before(40),
    )
    .unwrap();
    let noop = noop.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.package().to_bytes().unwrap(), exact);

    let mut changed = protected.edit();
    changed
        .patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(80),
        )
        .unwrap();
    assert!(matches!(
        changed.commit(),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::Protected
    ));

    for source in [
        Package::parse(r"{\rtf1\ansi{\stylesheet{\s1 Named;}}\s1 Styled}").unwrap(),
        Package::parse(r"{\rtf1\ansi\ls1 Listed}").unwrap(),
        Package::parse(r"{\rtf1\ansi One\line Two}").unwrap(),
        Package::parse(r"{\rtf1\ansi\trowd\cellx1000\intbl Cell\cell\row}").unwrap(),
    ] {
        let exact = source.to_bytes().unwrap();
        let mut edit = source.edit();
        assert!(matches!(
            edit.patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new().with_space_before(20)
            ),
            Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
        ));
        assert_eq!(edit.operation_count(), 0);
        assert_eq!(source.to_bytes().unwrap(), exact);
    }

    let raw = br"{\rtf1\ansi One}";
    let compressed =
        Package::from_bytes(&compress(raw, true, CompressionLimits::default()).unwrap()).unwrap();
    let mut edit = compressed.edit();
    assert!(matches!(
        edit.patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(20)
        ),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let mut non_ascii_bytes = br"{\rtf1\ansi\ansicpg1252 Caf".to_vec();
    non_ascii_bytes.push(0xe9);
    non_ascii_bytes.push(b'}');
    let non_ascii = Package::from_bytes(&non_ascii_bytes).unwrap();
    let mut edit = non_ascii.edit();
    assert!(matches!(
        edit.patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_space_before(20)
        ),
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedContent
    ));

    let source_text = r"{\rtf1 A}";
    let limits = SourceLimits::builder()
        .max_input_bytes(source_text.len())
        .expect("source-byte limit is representable")
        .build()
        .expect("source profile is consistent");
    let bounded =
        Package::parse_with_options(source_text, OpenOptions::new().source_limits(limits)).unwrap();
    let exact = bounded.to_bytes().unwrap();
    let mut edit = bounded.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_space_before(20),
    )
    .unwrap();
    assert!(matches!(edit.commit(), Err(Error::InputTooLarge { .. })));
    assert_eq!(bounded.to_bytes().unwrap(), exact);
}

#[test]
fn layout_refuses_each_exclusive_lifecycle_without_poisoning_it() {
    let source = Package::parse(r"{\rtf1\ansi First\par Second\par Third}").unwrap();
    for case in 0..4 {
        let mut edit = source.edit();
        match case {
            0 => {
                edit.remove_paragraph(Position::new(0)).unwrap();
            },
            1 => {
                edit.move_paragraph(Position::new(0), Position::new(2))
                    .unwrap();
            },
            2 => {
                edit.swap_paragraphs(Position::new(0), Position::new(2))
                    .unwrap();
            },
            _ => {
                edit.insert_paragraph(Position::new(0), "Added").unwrap();
            },
        }
        assert!(matches!(
            edit.patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new().with_space_before(40)
            ),
            Err(Error::Conflict {
                existing: 0,
                incoming: 1
            })
        ));
        assert_eq!(edit.operation_count(), 1);
        let commit = edit.commit().unwrap();
        assert_eq!(commit.diagnostics().operation_count(), 1);
        let expected = match case {
            0 => vec!["Second", "Third"],
            1 => vec!["Second", "Third", "First"],
            2 => vec!["Third", "Second", "First"],
            _ => vec!["Added", "First", "Second", "Third"],
        };
        let actual: Vec<_> = commit
            .package()
            .body()
            .paragraphs()
            .map(|p| p.to_text())
            .collect();
        assert_eq!(actual, expected);
    }
}

#[test]
fn layout_keeps_opaque_control_words_separate_from_reconstructed_text() {
    for source in [
        r"{\rtf1\ansi{\vendor\ql Alpha}}",
        r"{\rtf1\ansi{\vendor12\ql 34}}",
        r"{\rtf1\ansi{\vendor\ql -3}}",
        r"{\rtf1\ansi\vendor\ql Alpha}",
        r"{\rtf1\ansi\vendor12\ql 34}",
        r"{\rtf1\ansi\vendor\ql -3}",
        r"{\rtf1\ansi One\vendor\ql Two}",
        r"{\rtf1\ansi\vendor\other\ql Alpha}",
        r"{\rtf1\ansi\vendor  Alpha}",
        r"{\rtf1\ansi\|\ql Alpha}",
        r"{\rtf1\ansi{\*\vendor payload}Alpha}",
    ] {
        let source = Package::parse(source).unwrap();
        let mut edit = source.edit();
        edit.patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.package().text(), source.text());
        assert_eq!(commit.package().opaque(), source.opaque());
        assert_eq!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            source.to_bytes().unwrap()
        );
    }
}

#[test]
fn layout_preserves_character_coverage_when_unicode_encoding_splits_runs() {
    let source = Package::parse(r"{\rtf1\ansi\ansicpg1252 \'e9\'e9{\b \'e0\'e0}}").unwrap();
    let coverage = |p: &Package| {
        p.raw()
            .blocks()
            .iter()
            .flat_map(|b| {
                b.text
                    .chars()
                    .filter(|c| *c != '\n')
                    .map(move |c| (c, b.formatting))
            })
            .collect::<Vec<_>>()
    };
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.package().text(), source.text());
    assert_eq!(coverage(commit.package()), coverage(&source));
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
}

#[test]
fn layout_preserves_header_newlines_resources_and_trailer_destinations() {
    let input = concat!(
        "{\\rtf1\\ansi\r\n{\\fonttbl{\\f0\\fnil Native Font;}}\r\n",
        "{\\colortbl;\\red12\\green34\\blue56;}\r\n{\\info{\\title Keep title}}\r\n",
        "\\pard\\f0\\cf1{\\b One\\par}\r\n\\pard{\\i Two\\par}\r\n",
        "{\\*\\datastore}{\\*\\futuretail retained}}"
    );
    let source = Package::parse(input).unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.package().to_bytes().unwrap();
    let first = input.find(r"{\b One\par}").unwrap();
    let trailer = input.find(r"{\*\datastore}").unwrap();
    assert!(output.starts_with(&input.as_bytes()[..first]));
    assert!(output.ends_with(&input.as_bytes()[trailer..]));
    assert_eq!(commit.package().text(), source.text());
    assert_eq!(
        commit.package().raw().font_table().fonts(),
        source.raw().font_table().fonts()
    );
    assert_eq!(
        commit.package().raw().color_table().colors(),
        source.raw().color_table().colors()
    );
    assert_eq!(commit.package().raw().info(), source.raw().info());
    assert_eq!(
        commit.package().raw().data_store(),
        source.raw().data_store()
    );
    assert_eq!(commit.package().opaque(), source.opaque());
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        input.as_bytes()
    );
}

#[test]
fn layout_reemits_only_bookmark_events_inside_the_replaced_source_extent() {
    for text in [
        r"{\rtf1\ansi{\b {\*\bkmkstart bm}One{\*\bkmkend bm}\par}}",
        r"{\rtf1\ansi{\*\bkmkstart bm}{\b One}{\*\bkmkend bm}}",
        r"{\rtf1\ansi{\b {\*\bkmkstart bm}One}{\*\bkmkend bm}}",
        r"{\rtf1\ansi{\page\b One\softpage}}",
    ] {
        let source = Package::parse(text).unwrap();
        let mut edit = source.edit();
        edit.patch_paragraph_layout(
            Position::new(0),
            ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        assert_eq!(commit.package().text(), source.text());
        assert_eq!(
            commit.package().raw().body_story_events(),
            source.raw().body_story_events()
        );
        if let Some(bookmark) = commit.package().raw().bookmarks().bookmarks().first() {
            assert_eq!((bookmark.position, bookmark.content.as_ref()), (0, "One"));
            let output = String::from_utf8(commit.package().to_bytes().unwrap()).unwrap();
            assert_eq!(output.matches(r"\bkmkstart").count(), 1);
            assert_eq!(output.matches(r"\bkmkend").count(), 1);
        }
        assert_eq!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            text.as_bytes()
        );
    }
}

#[test]
fn interleaved_nonbody_metadata_is_refused_without_dropping_it() {
    let source = Package::parse(r"{\rtf1\ansi One{\colortbl;\red12\green34\blue56;}Two}").unwrap();
    let original = source.to_bytes().unwrap();
    let mut edit = source.edit();
    edit.patch_paragraph_layout(
        Position::new(0),
        ParagraphLayoutPatch::new().with_frame(ParagraphFrame::default()),
    )
    .unwrap();
    assert!(
        matches!(edit.commit(),Err(Error::Refused(r)) if r.reason()=="paragraph-layout source contains interleaved non-body destinations")
    );
    assert_eq!(source.to_bytes().unwrap(), original);
}
