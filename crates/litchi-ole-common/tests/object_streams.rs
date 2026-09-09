#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "integration tests use concise assertions and checked fixture-sized literals"
)]

use litchi_cfb::OleWriter;
use litchi_ole_common::object::{Editor, Limits, Target, Targets};
use litchi_ole_common::ole_streams::{
    CF_DIB, ClipboardFormat, Limits as StreamLimits, OleNativeStream, OlePresentationStream,
};
use std::io::Cursor;
use std::sync::Arc;

fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
    let mut writer = OleWriter::new();
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("test CFB should write");
    output.into_inner()
}

fn presentation(data: &[u8]) -> Vec<u8> {
    OlePresentationStream::new(ClipboardFormat::Standard(CF_DIB), data.to_vec())
        .expect("presentation should construct")
        .to_bytes()
        .expect("presentation should serialize")
}

fn native(data: &[u8]) -> Vec<u8> {
    OleNativeStream::new(data.to_vec())
        .expect("native stream should construct")
        .to_bytes()
        .expect("native stream should serialize")
}

fn targets() -> Targets {
    Targets::one(Target::new("object", ["ObjectPool", "_1"]).expect("target should validate"))
}

fn package() -> Vec<u8> {
    let first = presentation(b"first presentation");
    let second = presentation(b"second presentation");
    let native = native(b"opaque native bytes");
    write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_1"])
            .expect("object storage should write");
        writer
            .set_storage_metadata(
                &["ObjectPool", "_1"],
                0xA1A2_A3A4,
                0x0102_0304_0506_0708,
                0x1112_1314_1516_1718,
            )
            .expect("object metadata should write");
        writer
            .create_stream_with_metadata(
                &["ObjectPool", "_1", "\u{0002}OlePres042"],
                &second,
                0xB1B2_B3B4,
                0x2122_2324_2526_2728,
                0x3132_3334_3536_3738,
            )
            .expect("second presentation should write");
        writer
            .create_stream_with_metadata(
                &["ObjectPool", "_1", "\u{0002}OlePres007"],
                &first,
                0xC1C2_C3C4,
                0x4142_4344_4546_4748,
                0x5152_5354_5556_5758,
            )
            .expect("first presentation should write");
        writer
            .create_stream_with_metadata(
                &["ObjectPool", "_1", "\u{0001}Ole10Native"],
                &native,
                0xD1D2_D3D4,
                0x6162_6364_6566_6768,
                0x7172_7374_7576_7778,
            )
            .expect("native stream should write");
        writer
            .create_stream(&["ObjectPool", "_1", "Opaque"], b"untouched")
            .expect("opaque stream should write");
    })
}

fn metadata(object: &litchi_ole_common::object::Object, name: &str) -> (u32, u64, u64) {
    let directory = object
        .streams()
        .iter()
        .find(|stream| stream.name() == Some(name))
        .and_then(|stream| stream.directory())
        .expect("stream metadata should be captured");
    (
        directory.state_bits(),
        directory.creation_time(),
        directory.modified_time(),
    )
}

#[test]
fn object_and_snapshot_expose_typed_ole2_streams_without_name_scans() {
    let editor = Editor::open(package(), targets(), Limits::default()).expect("editor should open");
    let object = editor.objects().get("object").expect("object should exist");

    let presentations = object.presentations().expect("presentations should parse");
    assert_eq!(
        presentations
            .iter()
            .map(|(index, _)| *index)
            .collect::<Vec<_>>(),
        [7, 42]
    );
    assert_eq!(presentations[0].1.data(), b"first presentation");
    let captured = object
        .streams()
        .iter()
        .find(|stream| stream.path().len() == 1 && stream.name() == Some("\u{0002}OlePres007"))
        .expect("presentation stream should be captured");
    let direct = captured
        .presentation()
        .expect("direct presentation should parse")
        .expect("direct presentation should exist");
    assert!(Arc::ptr_eq(
        &direct.bytes_shared(),
        &captured.bytes_shared()
    ));
    assert!(Arc::ptr_eq(
        &presentations[0].1.bytes_shared(),
        &captured.bytes_shared()
    ));
    assert_eq!(
        object
            .presentation(42)
            .expect("indexed presentation should parse")
            .expect("indexed presentation should exist")
            .data(),
        b"second presentation"
    );
    assert_eq!(
        object
            .native()
            .expect("native stream should parse")
            .expect("native stream should exist")
            .data(),
        b"opaque native bytes"
    );

    let snapshot = editor.snapshot();
    assert_eq!(snapshot.presentations("object").unwrap().len(), 2);
    assert_eq!(
        snapshot.native("object").unwrap().unwrap().data(),
        b"opaque native bytes"
    );
}

#[test]
fn typed_owner_edits_are_noop_source_checked_and_reversible() {
    let original = package();
    let source_editor = Editor::open(original.clone(), targets(), Limits::default())
        .expect("source editor should open");
    let source = source_editor
        .snapshot()
        .presentation("object", 7)
        .expect("source presentation should parse")
        .expect("source presentation should exist");
    let mut typed_edit = source.edit();
    typed_edit
        .set_width(640)
        .expect("presentation edit should stage");
    let typed_commit = typed_edit
        .commit()
        .expect("presentation edit should commit");
    assert!(!typed_commit.patch().is_noop());
    assert_eq!(
        typed_commit
            .patch()
            .inverse()
            .apply(typed_commit.snapshot())
            .expect("inverse should apply")
            .bytes(),
        source.bytes()
    );

    let mut noop = Editor::open(original.clone(), targets(), Limits::default())
        .expect("no-op editor should open");
    noop.update_presentation("object", 7, |_transaction| Ok(()))
        .expect("no-op presentation edit should commit");
    assert!(!noop.is_changed());
    assert_eq!(noop.finish().expect("no-op should finish"), original);

    let mut owner =
        Editor::open(original.clone(), targets(), Limits::default()).expect("owner should open");
    owner
        .apply_presentation_patch("object", 7, typed_commit.patch())
        .expect("typed patch should publish");
    assert_eq!(
        owner
            .snapshot()
            .presentation("object", 7)
            .unwrap()
            .unwrap()
            .width(),
        640
    );
    owner
        .apply_presentation_patch("object", 7, &typed_commit.patch().inverse())
        .expect("inverse patch should publish through owner");
    assert_eq!(
        owner
            .snapshot()
            .presentation("object", 7)
            .unwrap()
            .unwrap()
            .bytes(),
        source.bytes()
    );

    let mut stale = Editor::open(original.clone(), targets(), Limits::default())
        .expect("stale editor should open");
    stale
        .update_presentation("object", 7, |transaction| {
            transaction.set_height(480)?;
            Ok(())
        })
        .expect("intervening edit should publish");
    let after_intervening = stale.snapshot().finish().expect("snapshot should finish");
    assert!(
        stale
            .apply_presentation_patch("object", 7, typed_commit.patch())
            .is_err()
    );
    assert_eq!(
        stale
            .snapshot()
            .finish()
            .expect("stale rejection should preserve owner"),
        after_intervening
    );
}

#[test]
fn native_owner_edit_and_explicit_limits_preserve_directory_metadata() {
    let original = package();
    let mut editor =
        Editor::open(original, targets(), Limits::default()).expect("editor should open");
    let object = editor.objects().get("object").expect("object should exist");
    let storage_metadata = (
        object.storage().directory().state_bits(),
        object.storage().directory().creation_time(),
        object.storage().directory().modified_time(),
    );
    let presentation_metadata = metadata(object, "\u{0002}OlePres007");
    let native_metadata = metadata(object, "\u{0001}Ole10Native");
    let opaque_metadata = metadata(object, "Opaque");
    let presentation = object
        .presentation(7)
        .unwrap()
        .expect("presentation should exist");
    let limits = StreamLimits {
        max_bytes: presentation.bytes().len(),
        max_data_bytes: presentation.data().len(),
        max_toc_entries: 999,
    };
    assert_eq!(
        editor
            .snapshot()
            .presentation_with_limits("object", 7, limits)
            .unwrap()
            .unwrap()
            .data(),
        b"first presentation"
    );
    editor
        .update_presentation("object", 7, |transaction| {
            transaction.set_width(640)?;
            Ok(())
        })
        .expect("presentation edit should publish");
    let native_source = editor
        .snapshot()
        .native("object")
        .unwrap()
        .expect("native source should exist");
    let mut native_transaction = native_source.edit();
    native_transaction
        .set_data(b"patched native bytes".to_vec())
        .expect("native patch should stage");
    let native_commit = native_transaction
        .commit()
        .expect("native patch should commit");
    editor
        .apply_native_patch("object", native_commit.patch())
        .expect("native patch should publish");
    editor
        .apply_native_patch("object", &native_commit.patch().inverse())
        .expect("native inverse should publish");
    editor
        .update_native("object", |transaction| {
            transaction.set_data(b"replacement native bytes".to_vec())?;
            Ok(())
        })
        .expect("native edit should publish");
    assert_eq!(
        editor.snapshot().native("object").unwrap().unwrap().data(),
        b"replacement native bytes"
    );

    let object = editor
        .objects()
        .get("object")
        .expect("object should remain");
    assert_eq!(
        (
            object.storage().directory().state_bits(),
            object.storage().directory().creation_time(),
            object.storage().directory().modified_time(),
        ),
        storage_metadata
    );
    assert_eq!(
        metadata(object, "\u{0002}OlePres007"),
        presentation_metadata
    );
    assert_eq!(metadata(object, "\u{0001}Ole10Native"), native_metadata);
    assert_eq!(metadata(object, "Opaque"), opaque_metadata);
}

#[test]
fn presentation_count_limit_is_independent_of_toc_entry_limit() {
    let stream = presentation(b"x");
    let package = write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_1"])
            .expect("object storage should write");
        for index in 0..1_000 {
            let name = format!("\u{0002}OlePres{index:03}");
            writer
                .create_stream(&["ObjectPool", "_1", &name], &stream)
                .expect("presentation stream should write");
        }
    });
    let editor = Editor::open(package, targets(), Limits::default()).expect("editor should open");
    let object = editor.objects().get("object").expect("object should exist");
    let limits = StreamLimits {
        max_bytes: stream.len(),
        max_data_bytes: 1,
        max_toc_entries: 1,
    };
    let error = object
        .presentations_with_limits(limits)
        .expect_err("a storage cannot expose 1,000 presentation streams");
    assert!(format!("{error:?}").contains("more than 999"));
    assert!(
        object.presentation(0).is_err(),
        "indexed access must enforce the storage count too"
    );
}

#[test]
fn absent_typed_streams_still_validate_explicit_limits() {
    let package = write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_1"])
            .expect("object storage should write");
        writer
            .create_stream(&["ObjectPool", "_1", "Opaque"], b"opaque")
            .expect("opaque stream should write");
    });
    let mut editor =
        Editor::open(package, targets(), Limits::default()).expect("editor should open");
    let invalid = StreamLimits {
        max_bytes: 0,
        max_data_bytes: 1,
        max_toc_entries: 1,
    };
    let object = editor.objects().get("object").expect("object should exist");
    let larger_toc_limit = StreamLimits {
        max_toc_entries: 1_000,
        ..StreamLimits::default()
    };
    assert!(object.presentations_with_limits(larger_toc_limit).is_ok());
    assert!(object.native_with_limits(larger_toc_limit).is_ok());
    let presentation_error = object
        .presentation_with_limits(7, invalid)
        .expect_err("absent presentation should still validate limits");
    assert!(format!("{presentation_error:?}").contains("limits must be non-zero"));
    let native_error = object
        .native_with_limits(invalid)
        .expect_err("absent native stream should still validate limits");
    assert!(format!("{native_error:?}").contains("limits must be non-zero"));
    let list_error = object
        .presentations_with_limits(invalid)
        .expect_err("empty presentation listing should still validate limits");
    assert!(format!("{list_error:?}").contains("limits must be non-zero"));
    assert!(
        editor
            .update_native_with_limits("object", invalid, |_transaction| Ok(()))
            .is_err()
    );
}

#[test]
fn malformed_known_stream_is_rejected_without_owner_publication() {
    let malformed = write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_1"])
            .expect("object storage should write");
        writer
            .create_stream(&["ObjectPool", "_1", "\u{0002}OlePres007"], &[1, 2, 3])
            .expect("malformed presentation should write");
    });
    let mut editor = Editor::open(malformed.clone(), targets(), Limits::default())
        .expect("opaque malformed package should open");
    assert!(editor.snapshot().presentation("object", 7).is_err());
    assert!(
        editor
            .update_presentation("object", 7, |_transaction| Ok(()))
            .is_err()
    );
    assert!(!editor.is_changed());
    assert_eq!(
        editor.finish().expect("rejected edit should finish"),
        malformed
    );
}
