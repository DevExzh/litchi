//! Exact native Word review metadata and inert, source-preserving comment edits.
use litchi_rtf::{AnnotationType, Package, RevisionType};

const FILE: &str = "word-review-controls.rtf";
const COMMENT: &str = "MODERN_COMMENT_003: review fixture comment.";
const REPLY: &str = "MODERN_REPLY_003: native reply retained.";
const BODY: &str = concat!(
    "Native Word review and references\n",
    "Fixture ID: word-review-controls-003\n",
    "REVIEW_ORIGINAL_003\n",
    "REVIEW_INSERT_003\n",
    "BOOKMARK_VALUE_003\n",
    "COMMENT_ANCHOR_003\n",
    "REFERENCE_FIELD_FOLLOWS\n",
    // The body coordinate space keeps ordinary field results in their own story.
    "\n",
    "REVIEW_CONTROLS_END\n",
);

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}

fn patch_limits() -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        2,
        8,
        256 * 1024,
        512 * 1024,
    )
}

fn review(package: &Package, comments: [&str; 2]) {
    assert_eq!(package.text(), BODY);
    assert_eq!(package.paragraph_count(), 9);
    let tree = package.raw();
    assert_eq!(tree.paragraph_count(), 9);
    assert_eq!(tree.paragraphs().len(), 9);
    let paragraphs = tree.paragraphs_with_content();
    assert_eq!(paragraphs.len(), 9);
    assert_eq!(paragraphs[7].text(), "");
    assert_eq!(paragraphs[8].text(), "REVIEW_CONTROLS_END");
    assert_eq!(tree.fields().len(), 1);
    assert_eq!(tree.fields()[0].instruction, " REF LITCHI_BOOKMARK_003 ");
    assert_eq!(tree.fields()[0].result, "BOOKMARK_VALUE_003");
    let field = &tree.fields()[0];
    assert_eq!(field.field_type, litchi_rtf::FieldType::Reference);
    assert_eq!(field.owner, litchi_rtf::FieldOwner::Body);
    assert_eq!(
        field.status,
        litchi_rtf::FieldStatus {
            edited: true,
            ..Default::default()
        }
    );
    let field_position =
        BODY.find("REFERENCE_FIELD_FOLLOWS\n").unwrap() + "REFERENCE_FIELD_FOLLOWS\n".len();
    assert_eq!(
        (field.position, field.range_end),
        (field_position, field_position)
    );
    let bookmarks = tree.bookmarks();
    assert_eq!(bookmarks.bookmarks().len(), 1);
    let bookmark = bookmarks.get("LITCHI_BOOKMARK_003").unwrap();
    assert_eq!(bookmark.position, BODY.find("BOOKMARK_VALUE_003").unwrap());
    assert_eq!(bookmark.content, "BOOKMARK_VALUE_003");
    assert_eq!(
        tree.revision_authors()
            .iter()
            .map(|author| author.name.as_ref())
            .collect::<Vec<_>>(),
        ["Unknown", "Ryker Zhu"]
    );
    let revisions = tree.revisions();
    assert_eq!(revisions.len(), 2);
    let start = BODY.find("REVIEW_INSERT_003").unwrap();
    let end = start + "REVIEW_INSERT_003".len();
    for (revision, kind, content, position, range_end) in [
        (
            &revisions[0],
            RevisionType::Insertion,
            "REVIEW_INSERT_003",
            start,
            end,
        ),
        (
            &revisions[1],
            RevisionType::Deletion,
            "REVIEW_DELETE_003",
            end,
            end,
        ),
    ] {
        assert_eq!(revision.revision_type, kind);
        assert_eq!(revision.content, content);
        assert_eq!(revision.id, 1);
        assert_eq!(revision.author, "Ryker Zhu");
        assert_eq!(revision.date.as_deref(), Some("-2014676111"));
        assert_eq!(
            (revision.position, revision.range_end),
            (position, range_end)
        );
    }
    let start = BODY.find("COMMENT_ANCHOR_003").unwrap();
    assert_eq!(tree.annotations().len(), 2);
    for (comment, id, date, parent, text) in [
        (
            &tree.annotations()[0],
            1_194_253_998,
            "-2014676110",
            None,
            comments[0],
        ),
        (
            &tree.annotations()[1],
            2_055_799_712,
            "-2014676109",
            Some("-1"),
            comments[1],
        ),
    ] {
        assert_eq!(comment.annotation_type, AnnotationType::Comment);
        assert_eq!(comment.id, id);
        assert!(comment.has_reference);
        assert_eq!(comment.author, "Ryker Zhu");
        assert_eq!(comment.initials, "RZ");
        assert_eq!(comment.date.as_deref(), Some(date));
        // Retain Word's literal parent value; do not invent a resolved thread identity.
        assert_eq!(comment.parent_id.as_deref(), parent);
        assert_eq!(comment.text, text);
        assert_eq!(
            (comment.position, comment.range_end),
            (start, start + "COMMENT_ANCHOR_003".len())
        );
        assert!(comment.icon.is_none() && comment.time.is_none());
        assert!(comment.shapes.is_empty() && comment.shape_groups.is_empty());
        assert!(comment.drawing_order.is_empty() && comment.story_events.is_empty());
    }
}

#[test]
fn native_review_comments_and_revisions_keep_exact_authors_dates_and_body_ranges() {
    let bytes = native(FILE);
    let source = Package::from_bytes(&bytes).unwrap();
    review(&source, [COMMENT, REPLY]);
    let mut encoded = Vec::new();
    litchi_rtf::write::Writer::new(&mut encoded)
        .write_document(source.raw())
        .unwrap();
    let reopened = Package::from_bytes(&encoded).unwrap();
    review(&reopened, [COMMENT, REPLY]);
    assert_eq!(reopened.raw().annotations(), source.raw().annotations());
    assert_eq!(reopened.raw().revisions(), source.raw().revisions());
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn native_comment_body_edits_preserve_review_metadata_and_every_other_source_byte() {
    let bytes = native(FILE);
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit();
    edit.set_annotation_text(0, "EDITED_ROOT_003").unwrap();
    edit.set_annotation_text(1, "EDITED_REPLY_003").unwrap();
    let commit = edit.commit().unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    commit.package().save(file.path()).unwrap();
    let output = std::fs::read(file.path()).unwrap();
    let original = std::str::from_utf8(&bytes).unwrap();
    assert_eq!(original.matches(COMMENT).count(), 1);
    assert_eq!(original.matches(REPLY).count(), 1);
    let expected = original
        .replace(COMMENT, "EDITED_ROOT_003")
        .replace(REPLY, "EDITED_REPLY_003");
    assert_eq!(output, expected.as_bytes());
    let reopened = Package::open(file.path()).unwrap();
    review(&reopened, ["EDITED_ROOT_003", "EDITED_REPLY_003"]);
    assert_eq!(reopened.raw().revisions(), source.raw().revisions());
    let wire = commit
        .patch()
        .to_durable(patch_limits())
        .unwrap()
        .to_deterministic_json()
        .unwrap();
    let durable =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &wire,
            patch_limits(),
        )
        .unwrap();
    assert_eq!(durable.to_deterministic_json().unwrap(), wire);
    assert_eq!(
        source
            .apply_durable(&durable)
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        output
    );
    assert_eq!(
        reopened
            .apply_durable(&durable.inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(native(FILE), bytes);
}

#[test]
fn native_comment_empty_and_unicode_edits_keep_identity_and_failed_staging_is_atomic() {
    use litchi_rtf::edit::Error;
    let bytes = native(FILE);
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit();
    edit.set_annotation_text(0, "").unwrap();
    assert!(matches!(
        edit.set_annotation_text(2, "missing"),
        Err(Error::DestinationOutOfRange("annotation"))
    ));
    assert!(matches!(
        edit.set_annotation_text(0, "duplicate"),
        Err(Error::Conflict {
            existing: 0,
            incoming: 1
        })
    ));
    let empty = edit.commit().unwrap();
    let output = empty.package().to_bytes().unwrap();
    assert_eq!(
        output,
        std::str::from_utf8(&bytes)
            .unwrap()
            .replace(COMMENT, "")
            .as_bytes()
    );
    let reopened = Package::from_bytes(&output).unwrap();
    review(&reopened, ["", REPLY]);
    let replacement = "{review} \\ café 中文 😀\tend";
    let mut edit = reopened.edit();
    edit.set_annotation_text(0, replacement).unwrap();
    let unicode = edit.commit().unwrap();
    let unicode_bytes = unicode.package().to_bytes().unwrap();
    let again = Package::from_bytes(&unicode_bytes).unwrap();
    review(&again, [replacement, REPLY]);
    let patch = unicode.patch().to_durable(patch_limits()).unwrap();
    assert!(matches!(
        source.apply_durable(&patch),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        again.apply_durable(&patch),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        again
            .apply_durable(&patch.inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        output
    );
    let restored = reopened.apply(&empty.patch().inverse()).unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), bytes);
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn native_review_body_run_replacement_has_an_explicit_source_profile_refusal() {
    use litchi_core::{Position, edit::RefusalKind};
    use litchi_rtf::edit::Error;
    let bytes = native(FILE);
    let source = Package::from_bytes(&bytes).unwrap();
    review(&source, [COMMENT, REPLY]);
    let mut edit = source.edit();
    edit.replace_body_run_texts(&[litchi_rtf::edit::RunTextReplacement::new(
        Position::new(0),
        Position::new(0),
        "Changed title",
    )])
    .unwrap();
    // The ordinary-run route excludes dependent review owners; it must not
    // publish a shortened body with stale annotation/revision offsets.
    assert!(matches!(edit.commit(), Err(Error::Refused(refusal))
        if refusal.kind() == RefusalKind::UnsupportedContent
        && refusal.reason() == "run-format editing refuses dependent body destinations or review content"));
    review(&source, [COMMENT, REPLY]);
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(native(FILE), bytes);
}

#[test]
fn native_tracked_move_revisions_keep_signed_dates_authors_and_visible_body_anchors() {
    const MOVED: &str =
        "SOURCE_018: Move this complete native paragraph to the destination below.\n";
    let body = concat!(
        "WORD_TRACKED_MOVES_018\n",
        "Fixture ID: word-tracked-moves-018\n",
        "ANCHOR_018: This paragraph remains in place.\n",
        "DESTINATION_018\n",
        "SOURCE_018: Move this complete native paragraph to the destination below.\n",
        "END_TRACKED_MOVES_018\n",
    );
    let bytes = native("word-tracked-moves.rtf");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut output = Vec::new();
    litchi_rtf::write::Writer::new(&mut output)
        .write_document(source.raw())
        .unwrap();
    let written = Package::from_bytes(&output).unwrap();
    for package in [&source, &written] {
        assert_eq!(package.text(), body);
        assert_eq!(package.paragraph_count(), 6);
        assert!(package.annotations().is_empty());
        let tree = package.raw();
        assert_eq!(
            tree.revision_authors()
                .iter()
                .map(|a| a.name.as_ref())
                .collect::<Vec<_>>(),
            ["Unknown", "Ryker Zhu"]
        );
        assert_eq!(tree.revisions().len(), 2);
        let deletion = &tree.revisions()[0];
        let insertion = &tree.revisions()[1];
        assert_eq!(deletion.revision_type, RevisionType::Deletion);
        assert_eq!(insertion.revision_type, RevisionType::Insertion);
        for revision in [deletion, insertion] {
            assert_eq!(revision.id, 1);
            assert_eq!(revision.author, "Ryker Zhu");
            assert_eq!(revision.date.as_deref(), Some("-2014675705"));
            assert_eq!(revision.content, MOVED);
        }
        let from = body.find("ANCHOR_018").unwrap();
        let to = body.find(MOVED).unwrap();
        assert_eq!((deletion.position, deletion.range_end), (from, from));
        assert_eq!(
            (insertion.position, insertion.range_end),
            (to, to + MOVED.len())
        );
        let marks = tree.move_bookmarks();
        assert_eq!(marks.len(), 2);
        assert_eq!(marks[0].position, from);
        assert_eq!(marks[1].position, to);
        assert!(marks[0].content.is_empty());
        assert_eq!(marks[1].content, MOVED);
        for mark in marks {
            assert_eq!(mark.author, 1);
            assert_eq!(mark.date, 0x87ea7d07);
            assert_eq!(mark.tag, "move242987239");
        }
    }
    assert_eq!(
        written.raw().move_bookmarks(),
        source.raw().move_bookmarks()
    );
    assert_eq!(written.raw().revisions(), source.raw().revisions());
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(native("word-tracked-moves.rtf"), bytes);
}
