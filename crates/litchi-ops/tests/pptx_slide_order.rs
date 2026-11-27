//! Unified slide ordering delegates preservation and proofs to the native owner.
#![cfg(feature = "pptx")]
use litchi_core::{FileFormat, Position};
use litchi_ops::{BatchItem, BatchMode, Code, File, ItemResult, NewNode, Placement, props};

fn fixture() -> File {
    let mut file = litchi_ops::create(FileFormat::Pptx).unwrap();
    for (i, text) in ["A", "B", "C"].into_iter().enumerate() {
        file = file
            .add("/", NewNode::new("slide").prop("layout", "Blank"))
            .unwrap()
            .file;
        file = file
            .add(
                &format!("/slide[{}]", i + 1),
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
        .map(|i| {
            file.get(&format!("/slide[{i}]/shape[1]"))
                .unwrap()
                .view()
                .text()
                .unwrap()
                .unwrap()
                .into_owned()
        })
        .collect()
}
#[test]
fn every_move_placement_resolves_against_original_siblings() {
    let source = fixture();
    let before = source.to_bytes().unwrap();
    for from in 0..3 {
        for to in 0..3 {
            let mut expected = vec!["A", "B", "C"];
            let moved = expected.remove(from);
            expected.insert(to, moved);
            let file = source
                .move_to(&format!("/slide[{}]", from + 1), Position::new(to))
                .unwrap();
            assert_eq!(texts(&file), expected);
        }
        for anchor in 0..3 {
            for after in [false, true] {
                let mut expected = vec!["A", "B", "C"];
                if from != anchor {
                    let moved = expected.remove(from);
                    expected.insert(
                        anchor - usize::from(from < anchor) + usize::from(after),
                        moved,
                    );
                }
                let path = format!("/slide[{}]", anchor + 1);
                let placement = if after {
                    Placement::after(&path)
                } else {
                    Placement::before(&path)
                }
                .unwrap();
                let file = source
                    .move_to(&format!("/slide[{}]", from + 1), placement)
                    .unwrap();
                assert_eq!(texts(&file), expected);
            }
        }
    }
    for (placement, expected) in [
        (Placement::Start, vec!["B", "A", "C"]),
        (Placement::End, vec!["A", "C", "B"]),
        (Placement::inside("/").unwrap(), vec!["A", "C", "B"]),
        (
            Placement::inside_at("/", Position::new(0)).unwrap(),
            vec!["B", "A", "C"],
        ),
    ] {
        assert_eq!(
            texts(&source.move_to("/slide[2]", placement).unwrap()),
            expected
        );
    }
    assert_eq!(source.to_bytes().unwrap(), before);
}
#[test]
fn ordered_edits_replay_undo_and_reopen() {
    let source = fixture();
    let before = source.to_bytes().unwrap();
    let commit = source
        .edit(|edit| {
            edit.move_to("/slide[1]", Placement::End)?;
            edit.swap("/slide[1]", "/slide[2]")?;
            edit.set("/slide[3]/shape[1]", props! { "text" => "updated A" })
        })
        .unwrap();
    assert_eq!(texts(commit.file()), ["C", "B", "updated A"]);
    let bytes = commit.file().to_bytes().unwrap();
    assert_eq!(
        commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
        bytes
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(
        texts(&litchi_ops::open_bytes(bytes).unwrap()),
        ["C", "B", "updated A"]
    );
    assert_eq!(source.to_bytes().unwrap(), before);
    assert_eq!(
        source
            .swap("/slide[1]", "/slide[1]")
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(
        source
            .move_to("/slide[1]", Placement::Start)
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(
        texts(&source.swap("/slide[1]", "/slide[3]").unwrap()),
        ["C", "B", "A"]
    );
}
#[test]
fn order_boundaries_batches_and_source_owner_keep_original_snapshot() {
    let source = fixture();
    let before = source.to_bytes().unwrap();
    assert!(source.move_to("/slide[1]", Position::new(3)).is_err());
    assert!(
        source
            .move_to("/slide[1]", Placement::inside("/slide[2]").unwrap())
            .is_err()
    );
    assert!(source.swap("/slide[1]", "/slide[2]/shape[1]").is_err());
    let eager = litchi_ops::open_bytes_with(
        before.clone(),
        &litchi_ops::OpenOptions::new().source_mode(litchi_ops::SourceMode::Eager),
    )
    .unwrap();
    assert_eq!(
        texts(&eager.move_to("/slide[1]", Placement::End).unwrap()),
        ["B", "C", "A"]
    );
    assert_eq!(
        texts(&eager.swap("/slide[1]", "/slide[3]").unwrap()),
        ["C", "B", "A"]
    );
    let deferred = litchi_ops::open_bytes(before.clone()).unwrap();
    let moved = deferred.move_to("/slide[1]", Placement::End).unwrap();
    assert_eq!(texts(&moved), ["B", "C", "A"]);
    let swapped = deferred.swap("/slide[1]", "/slide[3]").unwrap();
    assert_eq!(texts(&swapped), ["C", "B", "A"]);
    assert_eq!(deferred.to_bytes().unwrap(), before);

    let source_commit = deferred
        .edit(|edit| {
            edit.move_to("/slide[1]", Placement::End)?;
            edit.set("/slide[3]/shape[1]", props! { "text" => "updated A" })
        })
        .unwrap();
    assert_eq!(texts(source_commit.file()), ["B", "C", "updated A"]);
    let source_bytes = source_commit.file().to_bytes().unwrap();
    assert_eq!(
        texts(&litchi_ops::open_bytes(source_bytes.clone()).unwrap()),
        ["B", "C", "updated A"]
    );
    assert_eq!(
        texts(&source_commit.patch().apply(&deferred).unwrap()),
        ["B", "C", "updated A"]
    );
    assert_eq!(
        source_commit
            .patch()
            .inverse()
            .apply(source_commit.file())
            .unwrap()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(deferred.to_bytes().unwrap(), before);
    let commands = || {
        let mut first = BatchItem::new("move");
        first.path = Some("/slide[1]".into());
        first.after = Some("/slide[3]".into());
        let mut invalid = BatchItem::new("swap");
        invalid.path = Some("/slide[1]".into());
        invalid.path2 = Some("/slide[2]/shape[1]".into());
        let mut last = BatchItem::new("swap");
        last.path = Some("/slide[1]".into());
        last.path2 = Some("/slide[2]".into());
        [first, invalid, last]
    };
    assert_eq!(
        source
            .batch(commands(), BatchMode::Atomic)
            .unwrap_err()
            .item(),
        Some(1)
    );
    let result = source.batch(commands(), BatchMode::BestEffort).unwrap();
    assert!(matches!(&result.value()[1], ItemResult::Failed(e) if e.item()==Some(1)));
    assert_eq!(texts(result.file()), ["C", "B", "A"]);
    assert_eq!(source.to_bytes().unwrap(), before);

    let source_result = deferred.batch(commands(), BatchMode::BestEffort).unwrap();
    assert!(matches!(&source_result.value()[1], ItemResult::Failed(e) if e.item()==Some(1)));
    assert_eq!(texts(source_result.file()), ["C", "B", "A"]);
    assert_eq!(deferred.to_bytes().unwrap(), before);
}

#[test]
fn ordering_admits_distinct_targets_and_staged_actions_before_editing() {
    use litchi_ops::transaction::SemanticLimits;
    let source = fixture();
    let before = source.to_bytes().unwrap();
    let single = SemanticLimits::builder()
        .max_expanded_targets(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(single),
                |edit| edit.move_to("/slide[1]", Placement::End)
            )
            .is_ok()
    );
    assert!(
        source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(single),
                |edit| edit.swap("/slide[1]", "/slide[1]")
            )
            .is_ok()
    );
    assert_eq!(
        source
            .edit_with_options(
                litchi_ops::EditOptions::new().semantic_limits(single),
                |edit| edit.swap("/slide[1]", "/slide[2]")
            )
            .unwrap_err()
            .code(),
        Code::LimitExceeded
    );
    for (bytes, succeeds) in [
        (size_of::<usize>() * 2, true),
        (size_of::<usize>() * 2 - 1, false),
    ] {
        let limits = SemanticLimits::builder()
            .max_staged_bytes(bytes)
            .unwrap()
            .build()
            .unwrap();
        let result = source.edit_with_options(
            litchi_ops::EditOptions::new().semantic_limits(limits),
            |edit| edit.swap("/slide[1]", "/slide[2]"),
        );
        assert_eq!(result.is_ok(), succeeds);
        if !succeeds {
            assert_eq!(result.unwrap_err().code(), Code::LimitExceeded);
        }
    }
    let result = source.edit(|edit| {
        edit.move_to("/slide[1]", Placement::End)?;
        let _ = edit.swap("/slide[1]", "/slide[2]/shape[1]");
        Ok(())
    });
    assert_eq!(result.unwrap_err().item(), Some(1));
    assert_eq!(source.to_bytes().unwrap(), before);
}
