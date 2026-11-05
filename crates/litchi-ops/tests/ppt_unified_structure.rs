//! Unified structural operations for legacy PowerPoint presentations.
//!
//! These tests deliberately exercise the public [`litchi_ops::File`] surface
//! through both retained eager and positional owners.  Native edits are used
//! only as an oracle for the immutable publication bytes and ordering.
#![cfg(feature = "ppt")]

use std::io::Cursor;

use litchi_core::Position;
use litchi_core::selector::{Path as SelectorPath, Query, Scope};
use litchi_ops::transaction::SemanticLimits;
use litchi_ops::{BatchItem, BatchMode, Code, File, OpenOptions, Placement, SourceMode};
use litchi_ppt::writer::Builder;

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new();
    for (index, text) in ["A", "B", "C", "D"].into_iter().enumerate() {
        let slide = builder.add_slide().expect("slide");
        builder
            .set_slide_timing(slide, litchi_ppt::writer::SlideTiming::default())
            .expect("slide timing");
        builder
            .add_textbox(slide, 10, 10, 240, 40, text)
            .unwrap_or_else(|error| panic!("shape {index}: {error}"));
    }
    builder
        .build()
        .expect("build PPT fixture")
        .to_bytes()
        .expect("serialize PPT fixture")
}

fn notes_fixture() -> Vec<u8> {
    let mut builder = Builder::new();
    for (index, text) in ["A", "B", "C"].into_iter().enumerate() {
        let slide = builder.add_slide().expect("slide");
        builder
            .set_slide_timing(slide, litchi_ppt::writer::SlideTiming::default())
            .expect("slide timing");
        builder
            .add_textbox(slide, 10, 10, 240, 40, text)
            .unwrap_or_else(|error| panic!("shape {index}: {error}"));
        builder
            .set_slide_notes(slide, &format!("notes-{text}"))
            .expect("speaker notes");
    }
    builder
        .build()
        .expect("build notes fixture")
        .to_bytes()
        .expect("serialize notes fixture")
}

fn eager(bytes: &[u8]) -> File {
    litchi_ops::open_bytes_with(
        bytes.to_vec(),
        &OpenOptions::new().source_mode(SourceMode::Eager),
    )
    .expect("open eager PPT")
}

fn positional(bytes: &[u8]) -> (tempfile::TempDir, File) {
    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("source.ppt");
    std::fs::write(&path, bytes).expect("write PPT source");
    let file = litchi_ops::open_with(&path, &OpenOptions::new()).expect("open positional PPT");
    (directory, file)
}

fn each_owner(bytes: &[u8], mut test: impl FnMut(File)) {
    test(eager(bytes));
    let (_directory, file) = positional(bytes);
    test(file);
}

fn scoped_slides(selector: &str) -> Query<'_> {
    Query::parse(selector)
        .expect("slide query")
        .with_scope(Scope::Path(Box::new(
            SelectorPath::parse("/").expect("root scope"),
        )))
}

fn slide_texts(file: &File) -> Vec<String> {
    let mut texts = Vec::new();
    for position in 1.. {
        let path = format!("/slide[{position}]/shape[1]");
        let Ok(node) = file.get(&path) else {
            break;
        };
        texts.push(
            node.view()
                .text()
                .expect("shape text")
                .expect("text body")
                .into_owned(),
        );
    }
    texts
}

fn native_slide_texts(bytes: &[u8]) -> Vec<String> {
    let package = litchi_ppt::Package::from_bytes(bytes).expect("native PPT");
    package
        .presentation()
        .expect("native presentation")
        .slides()
        .expect("native slides")
        .into_iter()
        .map(|slide| slide.text().expect("native slide text").to_owned())
        .collect()
}

fn native_remove(bytes: &[u8], positions_descending: &[usize]) -> Vec<u8> {
    let package = litchi_ppt::Package::from_bytes(bytes).expect("native PPT");
    let mut edit = package.edit().expect("native edit");
    for &position in positions_descending {
        edit.remove_slide(Position::new(position))
            .expect("native remove");
    }
    edit.commit()
        .expect("native commit")
        .into_package()
        .to_bytes()
        .expect("native bytes")
}

fn native_move(bytes: &[u8], from: usize, destination: usize) -> Vec<u8> {
    let package = litchi_ppt::Package::from_bytes(bytes).expect("native PPT");
    let mut edit = package.edit().expect("native edit");
    edit.move_slide(Position::new(from), Position::new(destination))
        .expect("native move");
    edit.commit()
        .expect("native commit")
        .into_package()
        .to_bytes()
        .expect("native bytes")
}

fn native_swap(bytes: &[u8], left: usize, right: usize) -> Vec<u8> {
    if left == right {
        return bytes.to_vec();
    }
    let package = litchi_ppt::Package::from_bytes(bytes).expect("native PPT");
    let mut edit = package.edit().expect("native edit");
    let (first, second) = if left < right {
        (left, right)
    } else {
        (right, left)
    };
    edit.move_slide(Position::new(first), Position::new(second))
        .expect("native first swap move");
    edit.move_slide(Position::new(second - 1), Position::new(first))
        .expect("native second swap move");
    edit.commit()
        .expect("native commit")
        .into_package()
        .to_bytes()
        .expect("native bytes")
}

fn native_notes(bytes: &[u8]) -> Vec<Option<String>> {
    let package = litchi_ppt::Package::from_bytes(bytes).expect("native PPT");
    package
        .presentation()
        .expect("native presentation")
        .slides()
        .expect("native slides")
        .into_iter()
        .map(|slide| {
            slide
                .speaker_notes()
                .expect("native notes")
                .map(|notes| notes.text().expect("notes text").to_owned())
        })
        .collect()
}

fn batch_move(path: &'static str, after: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("move");
    item.path = Some(path.into());
    item.after = Some(after.into());
    item
}

fn batch_swap(left: &'static str, right: &'static str) -> BatchItem<'static> {
    let mut item = BatchItem::new("swap");
    item.path = Some(left.into());
    item.path2 = Some(right.into());
    item
}

fn signed_fixture() -> Vec<u8> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(fixture()), litchi_cfb::OleFileLimits::default())
            .expect("open PPT fixture");
    let document = ole
        .open_stream(&["PowerPoint Document"])
        .expect("PowerPoint Document stream");
    let current_user = ole
        .open_stream(&["Current User"])
        .expect("Current User stream");

    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .expect("copy document stream");
    writer
        .create_stream(&["Current User"], &current_user)
        .expect("copy Current User stream");
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .expect("add signature stream");
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).expect("write signed PPT");
    bytes.into_inner()
}

#[test]
fn remove_path_and_descending_query_match_native_and_preserve_both_sources() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let removed = source.remove("/slide[2]").expect("remove slide");
        let removed_bytes = removed.to_bytes().expect("removed bytes");
        assert_eq!(slide_texts(&removed), ["A", "C", "D"]);
        assert_eq!(removed_bytes, native_remove(&bytes, &[1]));
        assert_eq!(native_slide_texts(&removed_bytes), ["A", "C", "D"]);
        assert_eq!(source.to_bytes().expect("unchanged source"), before);

        let selected = scoped_slides("slide[1], slide[3]");
        let removed = source.remove(selected).expect("remove query matches");
        let removed_bytes = removed.to_bytes().expect("removed query bytes");
        assert_eq!(slide_texts(&removed), ["B", "D"]);
        assert_eq!(removed_bytes, native_remove(&bytes, &[2, 0]));
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn remove_all_slides_follows_the_native_legacy_contract() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let removed = source
            .remove(scoped_slides("slide"))
            .expect("legacy PPT permits an empty slide list");
        let removed_bytes = removed.to_bytes().expect("empty PPT bytes");
        assert!(slide_texts(&removed).is_empty());
        assert_eq!(removed_bytes, native_remove(&bytes, &[3, 2, 1, 0]));
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn slide_removal_keeps_unselected_notes_and_native_dependency_closure() {
    let bytes = notes_fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let removed = source.remove("/slide[2]").expect("remove noted slide");
        let removed_bytes = removed.to_bytes().expect("removed notes bytes");
        assert_eq!(slide_texts(&removed), ["A", "C"]);
        assert_eq!(removed_bytes, native_remove(&bytes, &[1]));
        assert_eq!(
            native_notes(&removed_bytes),
            [Some("notes-A".into()), Some("notes-C".into())]
        );
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn move_placements_match_native_order_and_leave_the_input_unchanged() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");

        let end = source
            .move_to("/slide[1]", Placement::End)
            .expect("move to end");
        let end_bytes = end.to_bytes().expect("move bytes");
        assert_eq!(slide_texts(&end), ["B", "C", "D", "A"]);
        assert_eq!(end_bytes, native_move(&bytes, 0, 3));
        let start = source.move_to("/slide[4]", Placement::Start).unwrap();
        assert_eq!(slide_texts(&start), ["D", "A", "B", "C"]);
        let inside = source
            .move_to(
                "/slide[4]",
                Placement::inside_at("/", Position::new(1)).unwrap(),
            )
            .unwrap();
        assert_eq!(slide_texts(&inside), ["A", "D", "B", "C"]);
        for placement in [
            Placement::before("/slide[2]").unwrap(),
            Placement::after("/slide[2]").unwrap(),
        ] {
            let noop = source
                .edit(|edit| edit.move_to("/slide[2]", placement))
                .unwrap();
            assert!(!noop.diagnostics().changed());
            assert!(noop.patch().is_noop());
            assert_eq!(noop.file().to_bytes().unwrap(), before);
        }

        let before_anchor = source
            .move_to("/slide[4]", Placement::before("/slide[2]").unwrap())
            .expect("move before anchor");
        assert_eq!(slide_texts(&before_anchor), ["A", "D", "B", "C"]);
        assert_eq!(
            before_anchor.to_bytes().expect("before bytes"),
            native_move(&bytes, 3, 1)
        );

        let after_anchor = source
            .move_to("/slide[1]", Placement::after("/slide[2]").unwrap())
            .expect("move after anchor");
        assert_eq!(slide_texts(&after_anchor), ["B", "A", "C", "D"]);
        assert_eq!(
            after_anchor.to_bytes().expect("after bytes"),
            native_move(&bytes, 0, 1)
        );
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn swap_matches_native_two_move_order_and_exact_noops() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let swapped = source.swap("/slide[1]", "/slide[4]").expect("swap slides");
        let swapped_bytes = swapped.to_bytes().expect("swap bytes");
        assert_eq!(slide_texts(&swapped), ["D", "B", "C", "A"]);
        assert_eq!(swapped_bytes, native_swap(&bytes, 0, 3));
        assert_eq!(native_slide_texts(&swapped_bytes), ["D", "B", "C", "A"]);

        for left in 0..4 {
            for right in 0..4 {
                let swapped = source
                    .swap(
                        &format!("/slide[{}]", left + 1),
                        &format!("/slide[{}]", right + 1),
                    )
                    .expect("swap either direction");
                let mut expected = ["A", "B", "C", "D"];
                expected.swap(left, right);
                assert_eq!(slide_texts(&swapped), expected);
                assert_eq!(
                    swapped.to_bytes().unwrap(),
                    native_swap(&bytes, left, right)
                );
            }
        }

        let move_noop = source
            .move_to("/slide[2]", Position::new(1))
            .expect("same-position move");
        assert_eq!(move_noop.to_bytes().expect("move no-op bytes"), before);
        let swap_noop = source
            .swap("/slide[2]", "/slide[2]")
            .expect("same-target swap");
        assert_eq!(swap_noop.to_bytes().expect("swap no-op bytes"), before);
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });
}

#[test]
fn structural_transaction_replays_inverts_and_batches_atomically() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        let commit = source
            .edit(|edit| {
                edit.remove("/slide[2]")?;
                edit.move_to("/slide[1]", Placement::End)?;
                edit.swap("/slide[1]", "/slide[2]")
            })
            .expect("structural transaction");
        assert_eq!(slide_texts(commit.file()), ["D", "C", "A"]);
        let committed = commit.file().to_bytes().expect("commit bytes");
        assert_eq!(
            commit.patch().apply(&source).unwrap().to_bytes().unwrap(),
            committed
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
        assert_eq!(source.to_bytes().expect("unchanged source"), before);

        let commands = || {
            [
                batch_move("/slide[1]", "/slide[3]"),
                batch_swap("/slide[1]", "/slide[99]"),
                batch_swap("/slide[1]", "/slide[4]"),
            ]
        };
        let atomic = source
            .batch(commands(), BatchMode::Atomic)
            .expect_err("atomic batch must roll back");
        assert_eq!(atomic.item(), Some(1));
        assert_eq!(source.to_bytes().expect("atomic source"), before);

        let best_effort = source
            .batch(commands(), BatchMode::BestEffort)
            .expect("best effort batch");
        assert!(matches!(
            &best_effort.value()[1],
            litchi_ops::ItemResult::Failed(error)
                if error.item() == Some(1) && error.code() == Code::PathNotFound
        ));
        assert_eq!(slide_texts(best_effort.file()), ["D", "C", "A", "B"]);
        assert_eq!(source.to_bytes().expect("best effort source"), before);
    });
}

#[test]
fn invalid_targets_and_native_protection_refuse_without_mutation() {
    let bytes = fixture();
    each_owner(&bytes, |source| {
        let before = source.to_bytes().expect("source bytes");
        assert_eq!(
            source
                .remove("/slide[99]")
                .expect_err("missing slide")
                .code(),
            Code::PathNotFound
        );
        assert!(
            source
                .move_to("/slide[1]", Placement::before("/slide[99]").unwrap())
                .is_err()
        );
        assert!(source.swap("/slide[1]", "/slide[1]/shape[1]").is_err());
        assert!(
            source
                .move_to("/slide[1]", Placement::inside("/slide[2]").unwrap())
                .is_err()
        );
        assert_eq!(
            source
                .move_to("/slide[1]", Position::new(4))
                .unwrap_err()
                .code(),
            Code::OutOfRange
        );

        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    });

    let signed = litchi_ops::open_bytes(signed_fixture()).expect("signed PPT");
    assert!(matches!(signed, File::Signed(_)));
    assert_eq!(
        signed
            .remove("/slide[1]")
            .expect_err("signed structural edit")
            .code(),
        Code::DocumentProtected
    );

    let released = litchi_ops::open_bytes_with(bytes, &OpenOptions::new().retain_source(false))
        .expect("released PPT");
    assert_eq!(
        released
            .move_to("/slide[1]", Placement::End)
            .expect_err("released structural edit")
            .code(),
        Code::UnsupportedCapability
    );
}

#[test]
fn structural_limits_reject_before_native_publication() {
    let source = eager(&fixture());
    let before = source.to_bytes().expect("source bytes");
    let selected = scoped_slides("slide[1], slide[2]");
    let limits = [
        SemanticLimits::builder()
            .max_expanded_targets(1)
            .expect("expanded limit")
            .build()
            .expect("limits"),
        SemanticLimits::builder()
            .max_staged_bytes(1)
            .expect("staged limit")
            .build()
            .expect("limits"),
        SemanticLimits::builder()
            .max_snapshot_bytes(1)
            .expect("snapshot limit")
            .build()
            .expect("limits"),
    ];
    for limits in limits {
        assert_eq!(
            source
                .edit_with_options(
                    litchi_ops::EditOptions::new().semantic_limits(limits),
                    |edit| edit.remove(selected.clone())
                )
                .expect_err("structural limit")
                .code(),
            Code::LimitExceeded
        );
        assert_eq!(source.to_bytes().expect("unchanged source"), before);
    }
}

#[cfg(any(unix, windows))]
#[test]
fn stale_positional_structural_edit_is_refused_before_publication() {
    use std::io::Write as _;

    let directory = tempfile::tempdir().expect("PPT source directory");
    let path = directory.path().join("stale.ppt");
    std::fs::write(&path, fixture()).expect("write source");
    let source = litchi_ops::open_with(&path, &OpenOptions::new()).expect("open source");
    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(&path)
        .expect("reopen source");
    changed.write_all(b"changed").expect("change source");
    changed.flush().expect("flush source");

    assert_eq!(
        source
            .remove("/slide[1]")
            .expect_err("stale structural edit")
            .code(),
        Code::IoError
    );
}
