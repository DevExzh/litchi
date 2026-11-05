#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "fixtures use checked, small literals"
)]

use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::object::{Limits as ObjectLimits, Target, Targets, discover};
use litchi_ole_common::ole_streams::{
    CF_DIB, CF_METAFILEPICT, ClipboardFormat, Limits, OleNativeStream, OlePresentationStream,
    TOC_SIGNATURE, TocEntry, native, parse_named_native, parse_named_presentation, presentation,
    presentation_index, presentation_name,
};
use litchi_ole_common::{Error, LimitKind};
use std::io::Cursor;
use std::sync::Arc;

fn push_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn standard_format(output: &mut Vec<u8>, value: u32) {
    push_u32(output, u32::MAX);
    push_u32(output, value);
}

fn toc_entry() -> Vec<u8> {
    let mut output = Vec::new();
    standard_format(&mut output, CF_METAFILEPICT);
    push_u32(&mut output, 0);
    push_u32(&mut output, 1);
    push_u32(&mut output, u32::MAX);
    push_u32(&mut output, 0x20);
    output.extend_from_slice(&[0x95, 0x74, 0, 0, 0xaa, 0x42, 0, 0, 0x16, 0, 0, 0]);
    push_u32(&mut output, 2);
    push_u32(&mut output, 0x18);
    output
}

fn presentation(with_toc: bool) -> Vec<u8> {
    let mut output = Vec::new();
    standard_format(&mut output, CF_DIB);
    push_u32(&mut output, 4);
    push_u32(&mut output, 1);
    push_u32(&mut output, u32::MAX);
    push_u32(&mut output, 2);
    push_u32(&mut output, 0xdead_beef);
    push_u32(&mut output, 0x7491);
    push_u32(&mut output, 0x42a7);
    push_u32(&mut output, 3);
    output.extend_from_slice(b"DIB");
    if with_toc {
        push_u32(&mut output, TOC_SIGNATURE);
        push_u32(&mut output, 1);
        output.extend_from_slice(&toc_entry());
        output.extend_from_slice(b"future-tail");
    }
    output
}

fn metafile_presentation() -> Vec<u8> {
    let mut output = Vec::new();
    standard_format(&mut output, CF_METAFILEPICT);
    push_u32(&mut output, 4);
    push_u32(&mut output, 1);
    push_u32(&mut output, u32::MAX);
    push_u32(&mut output, 2);
    push_u32(&mut output, 0);
    push_u32(&mut output, 100);
    push_u32(&mut output, 200);
    push_u32(&mut output, 2);
    output.extend_from_slice(b"MF");
    output.extend_from_slice(&[0xa5; 18]);
    output
}

fn native(data: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_u32(&mut output, data.len() as u32);
    output.extend_from_slice(data);
    output
}

#[test]
fn presentation_reads_toc_and_retains_unknown_tail() {
    let bytes = presentation(true);
    let source: Arc<[u8]> = bytes.clone().into();
    let snapshot = presentation::Snapshot::parse_shared(Arc::clone(&source), Limits::default())
        .expect("presentation should parse");
    assert_eq!(snapshot.clipboard_format().standard_id(), Some(CF_DIB));
    assert!(snapshot.target_device().is_empty());
    assert_eq!(snapshot.width(), 0x7491);
    assert_eq!(snapshot.height(), 0x42a7);
    assert_eq!(snapshot.data(), b"DIB");
    assert_eq!(snapshot.toc_signature(), Some(TOC_SIGNATURE));
    assert_eq!(snapshot.toc_count(), Some(1));
    assert_eq!(snapshot.toc_entries().len(), 1);
    assert_eq!(
        snapshot.toc_entries()[0].clipboard_format().standard_id(),
        Some(CF_METAFILEPICT)
    );
    assert_eq!(snapshot.unknown_tail(), b"future-tail");
    assert!(Arc::ptr_eq(&snapshot.bytes_shared(), &source));
    assert_eq!(snapshot.bytes(), bytes.as_slice());
}

#[test]
fn presentation_edit_reparses_and_inverse_restores_exact_source() {
    let source = presentation::Snapshot::parse(&presentation(true), Limits::default())
        .expect("source should parse");
    let source_bytes = source.bytes().to_vec();
    let mut edit = source.edit();
    edit.set_width(640).expect("width should stage");
    edit.update_toc_entry(0, |entry| {
        entry.set_advf(9);
        Ok(())
    })
    .expect("TOC entry should stage");
    let commit = edit.commit().expect("edit should commit");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.package().width(), 640);
    assert_eq!(commit.package().toc_entries()[0].advf(), 9);
    assert_eq!(commit.package().reserved1(), 0xdead_beef);
    assert_eq!(
        commit.package().toc_entries()[0].reserved1(),
        &[0x95, 0x74, 0, 0, 0xaa, 0x42, 0, 0, 0x16, 0, 0, 0]
    );
    assert_eq!(commit.package().toc_entries()[0].reserved2(), 0x18);
    assert_eq!(commit.package().unknown_tail(), b"future-tail");
    assert_eq!(commit.patch().before_bytes(), source_bytes.as_slice());
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.package())
        .expect("inverse should apply");
    assert_eq!(restored.bytes(), source_bytes.as_slice());
    let stale = presentation::Snapshot::parse(&presentation(true), Limits::default())
        .expect("same bytes parse");
    let mut stale_edit = stale.edit();
    stale_edit.set_height(2).expect("stale edit should stage");
    let stale_commit = stale_edit.commit().expect("stale edit should commit");
    assert!(commit.patch().apply(&stale_commit.into_package()).is_err());

    let mut atomic = source.edit();
    assert!(
        atomic
            .update(|stream| {
                stream.set_width(99);
                stream.set_clipboard_format(ClipboardFormat::None, Limits::default())
            })
            .is_err()
    );
    assert_eq!(atomic.stream().width(), source.width());
}

#[test]
fn presentation_noop_shares_source_and_metafile_reserved_bytes_are_typed() {
    let source: Arc<[u8]> = metafile_presentation().into();
    let snapshot = presentation::Snapshot::parse_shared(Arc::clone(&source), Limits::default())
        .expect("metafile presentation should parse");
    assert_eq!(snapshot.reserved2(), Some(&[0xa5; 18]));
    let commit = snapshot.edit().commit().expect("no-op should commit");
    assert!(!commit.diagnostics().changed());
    assert!(Arc::ptr_eq(&commit.package().bytes_shared(), &source));
    assert_eq!(commit.package().bytes(), source.as_ref());
}

#[test]
fn signed_dimensions_from_stream_and_reverted_edits_preserve_source_allocation() {
    let mut bytes = presentation(false);
    bytes[28..32].copy_from_slice(&(-640i32).to_le_bytes());
    bytes[32..36].copy_from_slice(&(480i32).to_le_bytes());
    let source: Arc<[u8]> = bytes.into();
    let stream = OlePresentationStream::parse_shared(Arc::clone(&source), Limits::default())
        .expect("signed dimensions should parse");
    let snapshot = presentation::Snapshot::from_stream(stream, Limits::default())
        .expect("clean parsed stream should become a snapshot without reparsing");
    assert_eq!(snapshot.width(), -640);
    assert_eq!(snapshot.height(), 480);
    assert!(Arc::ptr_eq(&snapshot.bytes_shared(), &source));

    let mut edit = snapshot.edit();
    edit.set_width(-320).expect("width should stage");
    edit.set_width(-640).expect("width should revert");
    edit.set_height(720).expect("height should stage");
    edit.set_height(480).expect("height should revert");
    assert!(!edit.clone().commit().unwrap().diagnostics().changed());
    let commit = edit.commit().expect("reverted edit should commit");
    assert!(!commit.diagnostics().changed());
    assert!(Arc::ptr_eq(&commit.package().bytes_shared(), &source));

    let native_source: Arc<[u8]> = native(b"source native").into();
    let native_stream =
        OleNativeStream::parse_shared(Arc::clone(&native_source), Limits::default())
            .expect("native source should parse");
    let native_snapshot = native::Snapshot::from_stream(native_stream, Limits::default())
        .expect("clean native stream should become a snapshot without reparsing");
    let mut native_edit = native_snapshot.edit();
    native_edit
        .set_data(b"temporary native".to_vec())
        .expect("native edit should stage");
    native_edit
        .set_data(b"source native".to_vec())
        .expect("native edit should revert");
    assert!(
        !native_edit
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );
    let native_commit = native_edit
        .commit()
        .expect("reverted native edit should commit");
    assert!(!native_commit.diagnostics().changed());
    assert!(Arc::ptr_eq(
        &native_commit.package().bytes_shared(),
        &native_source
    ));
}

#[test]
fn alternate_standard_clipboard_marker_survives_typed_rewrite() {
    let mut bytes = presentation(false);
    bytes[0..4].copy_from_slice(&0xffff_fffeu32.to_le_bytes());
    let source = presentation::Snapshot::parse(&bytes, Limits::default())
        .expect("alternate marker should parse");
    assert_eq!(source.clipboard_format_marker(), 0xffff_fffe);
    let mut edit = source.edit();
    edit.set_width(321).expect("width should stage");
    let commit = edit.commit().expect("alternate marker edit should commit");
    assert_eq!(
        &commit.package().bytes()[0..4],
        0xffff_fffeu32.to_le_bytes().as_slice()
    );

    let mut toc = toc_entry();
    toc[0..4].copy_from_slice(&0xffff_fffeu32.to_le_bytes());
    let entry =
        TocEntry::parse(&toc, Limits::default()).expect("alternate TOC marker should parse");
    assert_eq!(entry.clipboard_format_marker(), 0xffff_fffe);
    let mut edited = entry.clone();
    edited.set_advf(edited.advf() + 1);
    let encoded = edited
        .to_bytes(Limits::default())
        .expect("TOC edit should serialize");
    assert_eq!(&encoded[0..4], 0xffff_fffeu32.to_le_bytes().as_slice());
}

#[test]
fn presentation_limits_reject_oversized_encoded_fields_before_use() {
    let mut malformed = Vec::new();
    standard_format(&mut malformed, CF_DIB);
    push_u32(&mut malformed, 4);
    push_u32(&mut malformed, 0);
    push_u32(&mut malformed, 0);
    push_u32(&mut malformed, 0);
    push_u32(&mut malformed, 0);
    push_u32(&mut malformed, 0);
    push_u32(&mut malformed, u32::MAX);
    assert!(OlePresentationStream::parse(&malformed, Limits::default()).is_err());

    let mut bounded = presentation(false);
    let limits = Limits {
        max_bytes: 8,
        max_data_bytes: 8,
        max_toc_entries: 1,
    };
    assert!(matches!(
        presentation::Snapshot::parse(&bounded, limits),
        Err(Error::Limit {
            kind: LimitKind::StreamBytes,
            ..
        })
    ));
    bounded.truncate(8);
    assert!(OlePresentationStream::parse(&bounded, Limits::default()).is_err());
}

#[test]
fn presentation_edits_and_serialization_use_retained_limits() {
    let limits = Limits {
        max_bytes: 4096,
        max_data_bytes: 3,
        max_toc_entries: 1,
    };
    let snapshot = presentation::Snapshot::parse(&presentation(false), limits)
        .expect("fixture should fit the custom data limit");
    let mut edit = snapshot.edit();
    assert!(edit.set_data(b"1234".to_vec()).is_err());
    assert_eq!(edit.stream().data(), b"DIB");
    edit.set_data(b"xy".to_vec())
        .expect("replacement inside the source limit should stage");
    let commit = edit.commit().expect("bounded replacement should commit");
    assert_eq!(commit.package().data(), b"xy");

    let too_small = Limits {
        max_bytes: commit.package().bytes().len() - 1,
        max_data_bytes: 3,
        max_toc_entries: 1,
    };
    assert!(commit.package().stream().to_bytes(too_small).is_err());
}

#[test]
fn native_stream_edit_is_source_checked_and_bounded() {
    let source = native::Snapshot::parse(&native(b"opaque native bytes"), Limits::default())
        .expect("native parses");
    assert_eq!(source.data(), b"opaque native bytes");
    let mut edit = source.edit();
    edit.set_data(b"replacement native bytes".to_vec())
        .expect("native replacement should stage");
    let commit = edit.commit().expect("native replacement should commit");
    assert_eq!(commit.package().data(), b"replacement native bytes");
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.package())
            .expect("native inverse should apply")
            .data(),
        b"opaque native bytes"
    );
    assert!(OleNativeStream::parse(&[0xff, 0xff, 0xff, 0xff], Limits::default()).is_err());
    let strict = Limits {
        max_bytes: 8,
        max_data_bytes: 2,
        max_toc_entries: 1,
    };
    assert!(matches!(
        native::Snapshot::parse(&native(b"123"), strict),
        Err(Error::Limit {
            kind: LimitKind::StreamBytes,
            ..
        })
    ));
}

#[test]
fn toc_entry_round_trips_and_rejects_trailing_bytes() {
    let bytes = toc_entry();
    let entry = TocEntry::parse(&bytes, Limits::default()).expect("TOCENTRY should parse");
    assert_eq!(
        entry.clipboard_format().standard_id(),
        Some(CF_METAFILEPICT)
    );
    assert!(entry.target_device().is_empty());
    assert_eq!(entry.aspect(), 1);
    assert_eq!(entry.lindex(), u32::MAX);
    assert_eq!(entry.tymed(), 0x20);
    assert_eq!(entry.advf(), 2);
    assert_eq!(entry.reserved2(), 0x18);
    assert_eq!(entry.bytes(), bytes.as_slice());
    assert!(TocEntry::parse(&[bytes.as_slice(), b"tail"].concat(), Limits::default()).is_err());
}

#[test]
fn non_nani_toc_signature_requires_zero_count_and_preserves_tail() {
    let mut bytes = presentation(false);
    push_u32(&mut bytes, 0x1234_5678);
    push_u32(&mut bytes, 0);
    bytes.extend_from_slice(b"producer-tail");
    let snapshot = presentation::Snapshot::parse(&bytes, Limits::default())
        .expect("zero-count non-NANI should parse");
    assert_eq!(snapshot.toc_signature(), Some(0x1234_5678));
    assert_eq!(snapshot.toc_count(), Some(0));
    assert!(snapshot.toc_entries().is_empty());
    assert_eq!(snapshot.unknown_tail(), b"producer-tail");
    assert_eq!(
        snapshot.stream().to_bytes(Limits::default()).unwrap(),
        bytes
    );

    let mut nonzero = presentation(false);
    push_u32(&mut nonzero, 0x1234_5678);
    push_u32(&mut nonzero, 1);
    assert!(presentation::Snapshot::parse(&nonzero, Limits::default()).is_err());

    let mut missing_count = presentation(false);
    push_u32(&mut missing_count, 0x1234_5678);
    assert!(presentation::Snapshot::parse(&missing_count, Limits::default()).is_err());
}

#[test]
fn toc_insert_remove_authoring_sets_nani_and_matching_count() {
    let source = presentation::Snapshot::parse(&presentation(false), Limits::default())
        .expect("source should parse");
    let entry = TocEntry::new(
        ClipboardFormat::Standard(CF_METAFILEPICT),
        Limits::default(),
    )
    .expect("new TOC entry should be valid");
    let mut edit = source.edit();
    edit.insert_toc_entry(0, entry)
        .expect("insertion should create a NANI table");
    let inserted = edit.commit().expect("insertion should commit");
    assert_eq!(inserted.package().toc_signature(), Some(TOC_SIGNATURE));
    assert_eq!(inserted.package().toc_count(), Some(1));
    assert_eq!(inserted.package().toc_entries().len(), 1);

    let mut remove = inserted.package().edit();
    let removed = remove
        .remove_toc_entry(0)
        .expect("entry should be removable");
    assert_eq!(
        removed.clipboard_format().standard_id(),
        Some(CF_METAFILEPICT)
    );
    let removed_commit = remove.commit().expect("removal should commit");
    assert_eq!(
        removed_commit.package().toc_signature(),
        Some(TOC_SIGNATURE)
    );
    assert_eq!(removed_commit.package().toc_count(), Some(0));
    assert!(removed_commit.package().toc_entries().is_empty());
    assert!(
        OlePresentationStream::parse(removed_commit.package().bytes(), Limits::default()).is_ok()
    );

    let mut invalid = source.edit();
    assert!(
        invalid
            .insert_toc_entry(
                1,
                TocEntry::new(ClipboardFormat::Standard(CF_DIB), Limits::default()).unwrap(),
            )
            .is_err()
    );
    assert!(!invalid.clone().commit().unwrap().diagnostics().changed());
}

#[test]
fn toc_edits_switch_non_nani_sources_to_nani_but_keep_unknown_tail() {
    let mut bytes = presentation(false);
    push_u32(&mut bytes, 0x1234_5678);
    push_u32(&mut bytes, 0);
    bytes.extend_from_slice(b"producer-tail");
    let source =
        presentation::Snapshot::parse(&bytes, Limits::default()).expect("source should parse");
    let mut edit = source.edit();
    edit.insert_toc_entry(
        0,
        TocEntry::new(
            ClipboardFormat::Standard(CF_METAFILEPICT),
            Limits::default(),
        )
        .unwrap(),
    )
    .expect("insertion should authorize NANI entries");
    let commit = edit.commit().expect("insertion should commit");
    assert_eq!(commit.package().toc_signature(), Some(TOC_SIGNATURE));
    assert_eq!(commit.package().toc_count(), Some(1));
    assert_eq!(commit.package().unknown_tail(), b"producer-tail");

    let mut remove = commit.package().edit();
    remove
        .remove_toc_entry(0)
        .expect("entry should be removable");
    let removed = remove.commit().expect("removal should commit");
    assert_eq!(removed.package().toc_count(), Some(0));
    assert_eq!(removed.package().unknown_tail(), b"producer-tail");
}

#[test]
fn named_stream_validation_and_real_cfb_object_integration() {
    assert_eq!(presentation_index("\u{0002}OlePres123").unwrap(), 123);
    assert_eq!(presentation_index("\u{0002}OlePres999").unwrap(), 999);
    assert_eq!(presentation_name(7).unwrap(), "\u{0002}OlePres007");
    assert!(presentation_index("\u{0002}OlePres1000").is_err());
    assert!(presentation_index("\u{0002}OlePres0x0").is_err());

    let presentation = presentation(false);
    let native = native(b"native");
    let mut writer = OleWriter::new();
    writer
        .create_storage(&["ObjectPool", "_1"])
        .expect("storage should be created");
    writer
        .create_stream(&["ObjectPool", "_1", "\u{0002}OlePres000"], &presentation)
        .expect("presentation stream should be created");
    writer
        .create_stream(&["ObjectPool", "_1", "\u{0001}Ole10Native"], &native)
        .expect("native stream should be created");
    let mut package = Cursor::new(Vec::new());
    writer.write_to(&mut package).expect("CFB should write");
    let bytes = package.into_inner();

    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
        .expect("CFB should open");
    let targets = Targets::one(Target::new("object", ["ObjectPool", "_1"]).unwrap());
    let objects =
        discover(&mut ole, &targets, ObjectLimits::default()).expect("object should discover");
    let object = objects.get("object").expect("object should be selected");
    let presentation_bytes = object
        .streams()
        .iter()
        .find(|stream| stream.name() == Some("\u{0002}OlePres000"))
        .expect("presentation stream should exist")
        .bytes_shared();
    let native_bytes = object
        .streams()
        .iter()
        .find(|stream| stream.name() == Some("\u{0001}Ole10Native"))
        .expect("native stream should exist")
        .bytes_shared();
    let parsed_presentation =
        parse_named_presentation("\u{0002}OlePres000", presentation_bytes, Limits::default())
            .expect("containing CFB stream should parse");
    let parsed_native = parse_named_native("\u{0001}Ole10Native", native_bytes, Limits::default())
        .expect("containing CFB native stream should parse");
    assert_eq!(parsed_presentation.data(), b"DIB");
    assert_eq!(parsed_native.data(), b"native");
}

#[test]
fn registered_clipboard_format_is_bounded_and_lossless() {
    let format =
        ClipboardFormat::registered(b"CustomFormat\0".to_vec(), Limits::default()).unwrap();
    let stream =
        OlePresentationStream::new(format.clone(), b"payload".to_vec(), Limits::default()).unwrap();
    let parsed = OlePresentationStream::parse(stream.bytes(), Limits::default()).unwrap();
    assert_eq!(parsed.clipboard_format(), &format);
    assert!(ClipboardFormat::registered(b"not-terminated".to_vec(), Limits::default()).is_err());
}

#[test]
fn primary_registered_bound_is_separate_from_toc_registered_formats() {
    let mut registered_bytes = vec![b'R'; 600];
    registered_bytes.push(0);
    let format = ClipboardFormat::registered(registered_bytes, Limits::default())
        .expect("format should fit stream cap");
    let entry = TocEntry::new(format.clone(), Limits::default())
        .expect("TOCENTRY should accept the registered name");
    let entry_bytes = entry
        .to_bytes(Limits::default())
        .expect("TOCENTRY should serialize");
    assert_eq!(
        TocEntry::parse(&entry_bytes, Limits::default())
            .unwrap()
            .clipboard_format(),
        &format
    );

    let mut bytes = presentation(false);
    push_u32(&mut bytes, TOC_SIGNATURE);
    push_u32(&mut bytes, 1);
    bytes.extend_from_slice(&entry_bytes);
    let limits = Limits {
        max_bytes: 4096,
        max_data_bytes: 8,
        max_toc_entries: 1,
    };
    let parsed = presentation::Snapshot::parse(&bytes, limits)
        .expect("TOC registered name should use the caller stream bound");
    assert_eq!(parsed.toc_entries()[0].clipboard_format(), &format);
    assert!(
        OlePresentationStream::new(format.clone(), b"payload".to_vec(), Limits::default()).is_err()
    );
    let primary = presentation::Snapshot::parse(&presentation(false), Limits::default()).unwrap();
    let mut edit = primary.edit();
    assert!(edit.set_clipboard_format(format).is_err());
    assert!(!edit.clone().commit().unwrap().diagnostics().changed());
}

#[test]
fn toc_entry_count_uses_configured_limit_instead_of_presentation_name_limit() {
    let mut bytes = presentation(false);
    push_u32(&mut bytes, TOC_SIGNATURE);
    push_u32(&mut bytes, 1000);
    for _ in 0..1000 {
        bytes.extend_from_slice(&toc_entry());
    }
    let limits = Limits {
        max_bytes: 128 * 1024,
        max_data_bytes: 8,
        max_toc_entries: 1000,
    };
    let parsed = presentation::Snapshot::parse(&bytes, limits)
        .expect("configured TOC limit should allow 1000 entries");
    assert_eq!(parsed.toc_count(), Some(1000));
    assert_eq!(parsed.toc_entries().len(), 1000);

    let default_count_limit = Limits {
        max_toc_entries: 999,
        ..limits
    };
    assert!(presentation::Snapshot::parse(&bytes, default_count_limit).is_err());
    assert_eq!(presentation_name(999).unwrap(), "\u{0002}OlePres999");
}

#[test]
fn target_device_checks_offsets_and_complete_devmode_bounds() {
    let mut driver = vec![0; 8];
    driver[0..2].copy_from_slice(&8u16.to_le_bytes());
    driver.extend_from_slice(b"driver\0");
    assert_eq!(
        litchi_ole_common::ole_streams::TargetDevice::from_bytes(driver.clone(), Limits::default())
            .unwrap()
            .bytes(),
        driver.as_slice()
    );

    let mut unterminated = vec![0; 8];
    unterminated[0..2].copy_from_slice(&8u16.to_le_bytes());
    unterminated.extend_from_slice(b"driver");
    assert!(
        litchi_ole_common::ole_streams::TargetDevice::from_bytes(unterminated, Limits::default())
            .is_err()
    );

    let mut devmode = vec![0; 8 + 156];
    devmode[6..8].copy_from_slice(&8u16.to_le_bytes());
    devmode[8 + 68..8 + 70].copy_from_slice(&156u16.to_le_bytes());
    assert!(
        litchi_ole_common::ole_streams::TargetDevice::from_bytes(devmode, Limits::default())
            .is_ok()
    );

    let mut short_devmode = vec![0; 8 + 156];
    short_devmode[6..8].copy_from_slice(&8u16.to_le_bytes());
    short_devmode[8 + 68..8 + 70].copy_from_slice(&70u16.to_le_bytes());
    assert!(
        litchi_ole_common::ole_streams::TargetDevice::from_bytes(short_devmode, Limits::default())
            .is_err()
    );
}
