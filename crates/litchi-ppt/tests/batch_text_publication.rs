#![allow(
    clippy::expect_used,
    clippy::too_many_lines,
    clippy::unwrap_used,
    reason = "canonical integration assertions use panic-driven test fixtures"
)]

//! Integration coverage for grouped native shape-text publication.
//!
//! The fixture deliberately crosses the native 256-item boundary while
//! retaining an unrelated auto shape on every slide.  The assertions use the
//! public package/edit API, so they cover target traversal, publication,
//! reopen, and exact reversible patch ownership together.

use std::collections::BTreeMap;
use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_core::Position;
use litchi_ppt::Package;
use litchi_ppt::package::edit::ShapeTextReplacement;
use litchi_ppt::shapes::{Shape, ShapeEnum, ShapeType};
use litchi_ppt::text_edit::Target;

const MAX_BATCH: usize = 256;

fn write_fixture(slide_count: usize, textboxes_per_slide: usize) -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    for slide_index in 0..slide_count {
        let slide = writer.add_slide().expect("add fixture slide");
        for shape_index in 0..textboxes_per_slide {
            let x = 10 + i32::try_from(shape_index % 16).expect("shape x") * 260;
            let y = 10 + i32::try_from(shape_index / 16).expect("shape y") * 60;
            writer
                .add_textbox(
                    slide,
                    x,
                    y,
                    240,
                    40,
                    &format!("source-{slide_index:02}-{shape_index:03}"),
                )
                .expect("add fixture textbox");
        }
        // This shape is intentionally outside every batch target.  Its
        // identity and empty text prove that the grouped publication leaves
        // unrelated drawing members alone.
        writer
            .add_rectangle(slide, 5000, 5000, 720, 480)
            .expect("add unrelated fixture rectangle");
    }

    writer
        .build()
        .expect("build fixture")
        .to_bytes()
        .expect("serialize fixture")
}

fn package(slide_count: usize, textboxes_per_slide: usize) -> Package {
    Package::from_bytes(&write_fixture(slide_count, textboxes_per_slide))
        .expect("open fixture package")
}

fn shape_texts(package: &Package) -> Vec<Vec<String>> {
    package
        .presentation()
        .expect("read fixture presentation")
        .slides()
        .expect("read fixture slides")
        .into_iter()
        .map(|slide| {
            slide
                .shapes()
                .expect("read fixture shapes")
                .iter()
                .map(|shape| shape.text().expect("read shape text"))
                .collect()
        })
        .collect()
}

fn unrelated_shapes(package: &Package, textboxes_per_slide: usize) -> Vec<(u32, ShapeType)> {
    package
        .presentation()
        .expect("read fixture presentation")
        .slides()
        .expect("read fixture slides")
        .into_iter()
        .map(|slide| {
            let shapes = slide.shapes().expect("read fixture shapes");
            let shape = shapes
                .get(textboxes_per_slide)
                .expect("unrelated rectangle exists");
            let ShapeEnum::AutoShape(shape) = shape else {
                panic!("fixture rectangle parsed as an unexpected shape")
            };
            (shape.id(), shape.shape_type())
        })
        .collect()
}

fn stream_map(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("open CFB stream map");
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).expect("read CFB stream");
            (path, data)
        })
        .collect()
}

fn assert_unrelated_streams_preserved(before: &[u8], after: &[u8]) {
    let before_streams = stream_map(before);
    let after_streams = stream_map(after);
    assert_eq!(
        before_streams.keys().collect::<Vec<_>>(),
        after_streams.keys().collect::<Vec<_>>(),
        "grouped publication must preserve CFB stream topology"
    );
    for (path, data) in before_streams {
        if path.last().is_some_and(|name| {
            name.eq_ignore_ascii_case("PowerPoint Document")
                || name.eq_ignore_ascii_case("Current User")
        }) {
            continue;
        }
        assert_eq!(
            after_streams.get(&path),
            Some(&data),
            "unrelated stream changed: {}",
            path.join("/")
        );
    }
}

fn assert_exact_patch_round_trip(
    source: &Package,
    target: &Package,
    patch: &litchi_ppt::package::edit::Patch,
) {
    let source_bytes = source.to_bytes().expect("read source bytes");
    let target_bytes = target.to_bytes().expect("read target bytes");
    assert_eq!(patch.before(), source_bytes.as_slice());
    assert_eq!(patch.after(), target_bytes.as_slice());

    let applied = source
        .apply(patch)
        .expect("apply exact forward patch")
        .into_package();
    assert_eq!(
        applied.to_bytes().expect("read applied bytes"),
        target_bytes,
        "forward patch must publish the committed bytes"
    );

    let restored = target
        .apply(&patch.inverse())
        .expect("apply exact inverse patch")
        .into_package();
    assert_eq!(
        restored.to_bytes().expect("read restored bytes"),
        source_bytes,
        "inverse patch must restore the exact source artifact"
    );
}

#[test]
fn grouped_batch_publishes_exact_256_targets_across_slides() {
    let source = package(2, 128);
    let source_bytes = source.to_bytes().expect("read source bytes");
    let source_texts = shape_texts(&source);
    let source_unrelated = unrelated_shapes(&source, 128);

    let replacements = (0..2)
        .flat_map(|slide| {
            (0..128).map(move |shape| {
                ShapeTextReplacement::new(
                    Target::new(Position::new(slide), Position::new(shape)),
                    format!("target-{slide:02}-{shape:03}"),
                )
            })
        })
        .collect::<Vec<_>>();
    assert_eq!(replacements.len(), MAX_BATCH);

    let mut edit = source.edit().expect("start grouped edit");
    assert_eq!(
        edit.set_shape_texts(&replacements)
            .expect("publish exact-boundary grouped batch"),
        MAX_BATCH
    );
    assert_eq!(edit.operation_count(), MAX_BATCH);
    let commit = edit.commit().expect("commit exact-boundary grouped batch");
    let target = commit.package();
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().operation_count(), MAX_BATCH);

    let target_texts = shape_texts(target);
    assert_eq!(target_texts.len(), source_texts.len());
    for slide in 0..2 {
        assert_eq!(target_texts[slide].len(), 129);
        for (shape, text) in target_texts[slide].iter().enumerate().take(128) {
            assert_eq!(*text, format!("target-{slide:02}-{shape:03}"));
        }
        assert_eq!(target_texts[slide][128], source_texts[slide][128]);
    }
    assert_eq!(unrelated_shapes(target, 128), source_unrelated);

    let target_bytes = target.to_bytes().expect("read target bytes");
    let reopened = Package::from_bytes(&target_bytes).expect("reopen grouped target");
    assert_eq!(shape_texts(&reopened), target_texts);
    assert_exact_patch_round_trip(&source, target, commit.patch());
    assert_eq!(
        source.to_bytes().expect("source remains immutable"),
        source_bytes
    );
    assert_unrelated_streams_preserved(&source_bytes, &target_bytes);
}

#[test]
fn grouped_batch_chunks_allow_out_of_order_repeated_targets() {
    let source = package(2, 128);
    let source_bytes = source.to_bytes().expect("read source bytes");
    let source_unrelated = unrelated_shapes(&source, 128);

    // The first chunk walks one slide in reverse order.  The second walks the
    // other slide in reverse order and deliberately repeats its first chunk's
    // target across the chunk boundary.
    let first_chunk = (0..128)
        .rev()
        .map(|shape| {
            ShapeTextReplacement::new(
                Target::new(Position::new(0), Position::new(shape)),
                format!("first-{shape:03}"),
            )
        })
        .collect::<Vec<_>>();
    let mut second_chunk = vec![ShapeTextReplacement::new(
        Target::new(Position::new(0), Position::new(127)),
        "second-repeat-127",
    )];
    second_chunk.extend((0..127).rev().map(|shape| {
        ShapeTextReplacement::new(
            Target::new(Position::new(1), Position::new(shape)),
            format!("second-{shape:03}"),
        )
    }));
    assert_eq!(first_chunk.len(), MAX_BATCH / 2);
    assert_eq!(second_chunk.len(), MAX_BATCH / 2);

    let mut edit = source.edit().expect("start chunked grouped edit");
    assert_eq!(
        edit.set_shape_texts(&first_chunk)
            .expect("publish first grouped chunk"),
        first_chunk.len()
    );
    assert_eq!(
        edit.set_shape_texts(&second_chunk)
            .expect("publish repeated grouped chunk"),
        second_chunk.len()
    );
    assert_eq!(edit.operation_count(), MAX_BATCH);
    let commit = edit.commit().expect("commit chunked grouped edit");
    let target = commit.package();
    let texts = shape_texts(target);

    for (shape, (first, second)) in texts[0].iter().zip(&texts[1]).enumerate().take(127) {
        assert_eq!(*first, format!("first-{shape:03}"));
        assert_eq!(*second, format!("second-{shape:03}"));
    }
    assert_eq!(texts[0][127], "second-repeat-127");
    assert_eq!(texts[1][127], "source-01-127");
    assert_eq!(unrelated_shapes(target, 128), source_unrelated);

    let target_bytes = target.to_bytes().expect("read chunked target bytes");
    let reopened = Package::from_bytes(&target_bytes).expect("reopen chunked target");
    assert_eq!(shape_texts(&reopened), texts);
    assert_exact_patch_round_trip(&source, target, commit.patch());
    assert_eq!(
        source.to_bytes().expect("source remains immutable"),
        source_bytes
    );
    assert_unrelated_streams_preserved(&source_bytes, &target_bytes);
}

#[test]
fn duplicate_targets_are_refused_per_chunk_and_leave_edits_retryable() {
    let source = package(1, 6);
    let mut edit = source.edit().expect("start duplicate-check edit");
    let first = Target::new(Position::new(0), Position::new(0));
    let second = Target::new(Position::new(0), Position::new(1));
    let third = Target::new(Position::new(0), Position::new(2));
    let fourth = Target::new(Position::new(0), Position::new(3));

    assert!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::new(first, "bad-first"),
            ShapeTextReplacement::new(first, "bad-duplicate"),
        ])
        .is_err(),
        "duplicate target in the first chunk must be refused"
    );
    assert_eq!(edit.operation_count(), 0);
    assert_eq!(shape_texts(edit.source())[0][0], "source-00-000");

    assert_eq!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::new(first, "good-first"),
            ShapeTextReplacement::new(second, "good-second"),
        ])
        .expect("retry first chunk after duplicate refusal"),
        2
    );
    assert_eq!(edit.operation_count(), 2);

    assert!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::new(third, "bad-third"),
            ShapeTextReplacement::new(third, "bad-duplicate-third"),
        ])
        .is_err(),
        "duplicate target in the second chunk must be refused"
    );
    assert_eq!(edit.operation_count(), 2);
    assert_eq!(shape_texts(edit.source())[0][2], "source-00-002");

    assert_eq!(
        edit.set_shape_texts(&[
            ShapeTextReplacement::new(third, "good-third"),
            ShapeTextReplacement::new(fourth, "good-fourth"),
        ])
        .expect("retry second chunk after duplicate refusal"),
        2
    );
    let commit = edit.commit().expect("commit retried chunks");
    let texts = shape_texts(commit.package());
    assert_eq!(texts[0][0], "good-first");
    assert_eq!(texts[0][1], "good-second");
    assert_eq!(texts[0][2], "good-third");
    assert_eq!(texts[0][3], "good-fourth");
    assert_eq!(texts[0][4], "source-00-004");
    assert_eq!(commit.patch().operation_count(), 4);
}

#[test]
fn same_text_noops_and_changed_mixed_batches_have_exact_publication_semantics() {
    let source = package(1, 4);
    let source_bytes = source.to_bytes().expect("read source bytes");
    let shared = source.shared_ole_file_arc();

    let noops = (0..4)
        .map(|shape| {
            ShapeTextReplacement::new(
                Target::new(Position::new(0), Position::new(shape)),
                format!("source-00-{shape:03}"),
            )
        })
        .collect::<Vec<_>>();
    let mut noop_edit = source.edit().expect("start no-op edit");
    assert_eq!(
        noop_edit
            .set_shape_texts(&noops)
            .expect("stage no-op batch"),
        0
    );
    assert_eq!(noop_edit.operation_count(), noops.len());
    let noop_commit = noop_edit.commit().expect("commit no-op batch");
    assert!(noop_commit.patch().is_noop());
    assert!(Arc::ptr_eq(
        &shared,
        &noop_commit.package().shared_ole_file_arc()
    ));
    assert_eq!(
        noop_commit.package().to_bytes().expect("read no-op bytes"),
        source_bytes
    );

    let mixed = [
        ShapeTextReplacement::new(
            Target::new(Position::new(0), Position::new(0)),
            "source-00-000",
        ),
        ShapeTextReplacement::new(
            Target::new(Position::new(0), Position::new(1)),
            "changed-001",
        ),
        ShapeTextReplacement::new(
            Target::new(Position::new(0), Position::new(2)),
            "source-00-002",
        ),
    ];
    let mut mixed_edit = source.edit().expect("start mixed edit");
    assert_eq!(
        mixed_edit
            .set_shape_texts(&mixed)
            .expect("stage mixed batch"),
        1
    );
    assert_eq!(mixed_edit.operation_count(), mixed.len());
    let mixed_commit = mixed_edit.commit().expect("commit mixed batch");
    assert!(!mixed_commit.patch().is_noop());
    assert_eq!(shape_texts(mixed_commit.package())[0][1], "changed-001");
    assert_eq!(shape_texts(mixed_commit.package())[0][0], "source-00-000");
    assert_eq!(shape_texts(mixed_commit.package())[0][2], "source-00-002");

    let target_bytes = mixed_commit
        .package()
        .to_bytes()
        .expect("read mixed target bytes");
    let reopened = Package::from_bytes(&target_bytes).expect("reopen mixed target");
    assert_eq!(shape_texts(&reopened), shape_texts(mixed_commit.package()));
    assert_exact_patch_round_trip(&source, mixed_commit.package(), mixed_commit.patch());
    assert_eq!(
        source.to_bytes().expect("source remains immutable"),
        source_bytes
    );
    assert_unrelated_streams_preserved(&source_bytes, &target_bytes);
}
