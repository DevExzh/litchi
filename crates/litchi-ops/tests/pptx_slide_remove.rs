//! Unified eager PPTX slide deletion delegates graph ownership to the native owner.
#![cfg(feature = "pptx")]

use litchi_core::FileFormat;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{
    BatchItem, BatchMode, Code, File, ItemResult, NewNode, OpenOptions, SourceMode, props,
};

fn fixture() -> File {
    let mut file = litchi_ops::create(FileFormat::Pptx).unwrap();
    for (index, text) in ["A", "B", "C"].into_iter().enumerate() {
        file = file
            .add("/", NewNode::new("slide").prop("layout", "Blank"))
            .unwrap()
            .file;
        file = file
            .add(
                &format!("/slide[{}]", index + 1),
                NewNode::new("shape")
                    .prop("text", text)
                    .prop("width", "1in")
                    .prop("height", "1in"),
            )
            .unwrap()
            .file;
    }
    file
}

fn texts(file: &File) -> Vec<String> {
    (1..=3)
        .filter_map(|index| {
            let path = format!("/slide[{index}]/shape[1]");
            let node = file.get(&path).ok()?;
            Some(
                node.view()
                    .text()
                    .expect("shape text")
                    .expect("text body")
                    .into_owned(),
            )
        })
        .collect()
}

fn scoped_slide_query(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("parse slide selector")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root selector scope"),
        )))
}

fn batch_remove(path: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("remove");
    item.path = Some(path.into());
    item
}

fn notes_fixture() -> Vec<u8> {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        let first = presentation.add_slide().unwrap();
        first.add_text_box("A", 0, 0, 914_400, 914_400);
        first.set_notes("speaker notes");
        presentation
            .add_slide()
            .unwrap()
            .add_text_box("B", 0, 0, 914_400, 914_400);
        presentation
            .add_slide()
            .unwrap()
            .add_text_box("C", 0, 0, 914_400, 914_400);
    }
    builder.build().unwrap().to_bytes().unwrap()
}

#[test]
fn path_removal_reindexes_slides_and_refuses_the_final_slide_atomically() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();

    let removed = source.remove("/slide[2]").expect("remove middle slide");
    assert_eq!(texts(&removed), ["A", "C"]);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let one_left = removed.remove("/slide[1]").expect("remove first slide");
    assert_eq!(texts(&one_left), ["C"]);
    let one_left_bytes = one_left.to_bytes().unwrap();
    let error = one_left
        .remove("/slide[1]")
        .expect_err("the final slide must be retained");
    assert_eq!(error.code(), Code::InvalidValue);
    assert_eq!(one_left.to_bytes().unwrap(), one_left_bytes);
}

#[test]
fn scoped_union_removes_descending_positions_and_all_slide_removal_is_atomic() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();

    let selected = scoped_slide_query("slide[1], slide[2]");
    assert_eq!(source.query(selected.clone()).unwrap().count(), 2);
    let commit = source
        .edit(|edit| edit.remove(selected))
        .expect("remove two selected slides");
    assert_eq!(texts(commit.file()), ["C"]);
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        original_bytes
    );
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let all = source
        .edit(|edit| edit.remove(scoped_slide_query("slide")))
        .expect_err("removing every slide must fail at the final slide");
    assert_eq!(all.code(), Code::InvalidValue);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);
}

#[test]
fn removal_rejects_zero_mixed_and_unsupported_targets_without_mutation() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();

    let zero = source
        .remove(scoped_slide_query("slide[99]"))
        .expect_err("zero-match query");
    assert_eq!(zero.code(), Code::ZeroMatches);

    let mixed = source
        .remove(scoped_slide_query("slide[1], slide[1] > shape[1]"))
        .expect_err("mixed slide and shape query");
    assert_eq!(mixed.code(), Code::InvalidArgument);

    let predicate = source
        .remove(scoped_slide_query("slide[unknownPredicate=true]"))
        .expect_err("unknown slide predicate");
    assert_eq!(predicate.code(), Code::InvalidSelector);

    let bare = source
        .remove("slide")
        .expect_err("bare selectors need a document-wide scope");
    assert_eq!(bare.code(), Code::BareSelectorRejected);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);
}

#[test]
fn source_backed_removal_survives_reopen_and_eager_removal_keeps_its_route() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();
    let reopened_source = litchi_ops::open_bytes(original_bytes.clone()).unwrap();
    let source_removed = reopened_source
        .remove("/slide[2]")
        .expect("source-backed removal");
    assert_eq!(texts(&source_removed), ["A", "C"]);
    assert_eq!(reopened_source.to_bytes().unwrap(), original_bytes);
    let source_one_left = source_removed
        .remove("/slide[1]")
        .expect("source-backed second removal");
    assert_eq!(texts(&source_one_left), ["C"]);
    let source_one_left_bytes = source_one_left.to_bytes().unwrap();
    let final_error = source_one_left
        .remove("/slide[1]")
        .expect_err("source-backed final slide refusal");
    assert_eq!(final_error.code(), Code::InvalidValue);
    assert_eq!(source_one_left.to_bytes().unwrap(), source_one_left_bytes);

    let eager = litchi_ops::open_bytes_with(
        original_bytes,
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .unwrap();
    let removed = eager.remove("/slide[2]").expect("eager removal");
    assert_eq!(texts(&removed), ["A", "C"]);

    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("removed.pptx");
    removed.save_as(&path).unwrap();
    let reopened = litchi_ops::open(&path).unwrap();
    assert_eq!(texts(&reopened), ["A", "C"]);
}

#[test]
fn source_backed_query_transaction_batch_and_inverse_are_failure_atomic() {
    let source = litchi_ops::open_bytes(fixture().to_bytes().unwrap()).unwrap();
    let original_bytes = source.to_bytes().unwrap();

    let query = scoped_slide_query("slide[1], slide[2]");
    let commit = source
        .edit(|edit| edit.remove(query))
        .expect("source query removal");
    assert_eq!(texts(commit.file()), ["C"]);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        original_bytes
    );

    let commands = || [batch_remove("/slide[2]"), batch_remove("/slide[99]")];
    let atomic = source
        .batch(commands(), BatchMode::Atomic)
        .expect_err("source atomic batch must roll back");
    assert_eq!(atomic.item(), Some(1));
    assert_eq!(atomic.code(), Code::PathNotFound);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let best_effort = source
        .batch(commands(), BatchMode::BestEffort)
        .expect("source best-effort batch");
    assert!(matches!(
        &best_effort.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &best_effort.value()[1],
        ItemResult::Failed(error) if error.item() == Some(1) && error.code() == Code::PathNotFound
    ));
    assert_eq!(texts(best_effort.file()), ["A", "C"]);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);
}

#[test]
fn removal_patch_replays_inverts_reopens_and_allows_following_edits() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| edit.remove("/slide[2]"))
        .expect("remove middle slide in one edit");
    assert_eq!(texts(commit.file()), ["A", "C"]);

    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        commit.file().to_bytes().unwrap()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        original_bytes
    );
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let edited = commit
        .file()
        .set("/slide[2]/shape[1]", props! { "text" => "updated C" })
        .expect("edit a surviving slide");
    assert_eq!(texts(&edited), ["A", "updated C"]);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("reopened.pptx");
    edited.save_as(&path).unwrap();
    assert_eq!(texts(&litchi_ops::open(&path).unwrap()), ["A", "updated C"]);
}

#[test]
fn removal_honors_expanded_target_and_staged_byte_limits() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();

    let target_limit = SemanticLimits::builder()
        .max_expanded_targets(1)
        .unwrap()
        .build()
        .unwrap();
    let target_error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(target_limit),
            |edit| edit.remove(scoped_slide_query("slide[1], slide[2]")),
        )
        .expect_err("two slides exceed the target bound");
    assert_eq!(target_error.code(), Code::LimitExceeded);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let exact_limit = SemanticLimits::builder()
        .max_staged_bytes(2 * size_of::<(usize, usize, usize)>())
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(exact_limit),
                |edit| edit.remove(scoped_slide_query("slide[1], slide[2]"))
            )
            .is_ok()
    );

    let staged_limit = SemanticLimits::builder()
        .max_staged_bytes(size_of::<usize>() - 1)
        .unwrap()
        .build()
        .unwrap();
    let staged_error = source
        .edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(staged_limit),
            |edit| edit.remove("/slide[2]"),
        )
        .expect_err("one removal exceeds the staged action bound");
    assert_eq!(staged_error.code(), Code::LimitExceeded);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);
}

#[test]
fn removal_batches_are_atomic_or_best_effort_at_item_boundaries() {
    let source = fixture();
    let original_bytes = source.to_bytes().unwrap();
    let commands = || {
        [
            batch_remove("/slide[2]"),
            batch_remove("/slide[99]"),
            batch_remove("/slide[2]"),
        ]
    };

    let atomic = source
        .batch(commands(), BatchMode::Atomic)
        .expect_err("atomic removal must roll back after the missing path");
    assert_eq!(atomic.item(), Some(1));
    assert_eq!(atomic.code(), Code::PathNotFound);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);

    let best_effort = source
        .batch(commands(), BatchMode::BestEffort)
        .expect("best-effort removal");
    assert!(matches!(
        &best_effort.value()[0],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert!(matches!(
        &best_effort.value()[1],
        ItemResult::Failed(error) if error.item() == Some(1) && error.code() == Code::PathNotFound
    ));
    assert!(matches!(
        &best_effort.value()[2],
        ItemResult::Applied(applied) if applied.changed
    ));
    assert_eq!(texts(best_effort.file()), ["A"]);
    assert_eq!(source.to_bytes().unwrap(), original_bytes);
}

#[test]
fn native_final_slide_refuses_and_notes_slides_leave_with_their_slide() {
    let mut builder = litchi_pptx::builder::Builder::new().unwrap();
    builder.presentation().unwrap().add_slide().unwrap();
    let package = builder.build().unwrap();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    assert!(edit.remove_slide(litchi_core::Position::new(0)).is_err());
    let (after, patch) = edit.commit().unwrap().into_parts();
    assert!(patch.is_noop());
    assert_eq!(after.to_bytes().unwrap(), before);

    // Record 0303: a slide's own notes slide leaves with it, on the eager
    // and the source-backed owner.
    for mode in [SourceMode::Eager, SourceMode::Auto] {
        let dependency =
            litchi_ops::open_bytes_with(notes_fixture(), &OpenOptions::new().source_mode(mode))
                .unwrap();
        let removed = dependency.remove("/slide[1]").unwrap();
        assert_eq!(texts(&removed), ["B", "C"]);
        // Removing every slide is still refused atomically: the final-slide
        // rule fires after two removals were staged.
        let before = dependency.to_bytes().unwrap();
        let late = dependency
            .remove(scoped_slide_query("slide[1], slide[2], slide[3]"))
            .unwrap_err();
        assert_eq!(late.code(), Code::InvalidValue);
        assert_eq!(dependency.to_bytes().unwrap(), before);
    }
}

#[test]
fn source_removal_keeps_retained_slide_notes_and_late_refusal_is_atomic() {
    let source = litchi_ops::open_bytes(notes_fixture()).unwrap();
    let original = source.to_bytes().unwrap();
    let removed = source.remove("/slide[2]").unwrap();
    assert_eq!(texts(&removed), ["A", "C"]);
    assert_eq!(
        removed
            .get("/slide[1]/notes")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap(),
        "speaker notes"
    );
    let reordered = removed
        .move_to("/slide[1]", litchi_ops::Placement::End)
        .unwrap();
    assert_eq!(texts(&reordered), ["C", "A"]);
    assert_eq!(
        reordered
            .get("/slide[2]/notes")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .unwrap(),
        "speaker notes"
    );
    let failure = source
        .remove(scoped_slide_query("slide[1], slide[2], slide[3]"))
        .unwrap_err();
    assert_eq!(failure.code(), Code::InvalidValue);
    assert_eq!(source.to_bytes().unwrap(), original);
}
