//! Regression tests for the DOC embedded-object model and codec.

use super::Limits;
use super::codec::{discover_targets, validate_existing_fields};
use super::model::{FieldMarker, Info, Kind};
use super::storage::{OBJECT_POOL, is_object_storage_name};
use super::{Snapshot, TransactionError};
use crate::writer::Writer;
use litchi_cfb::OleWriter;
use litchi_ole_common::object::{Editor as ObjectEditor, Target, Targets};
use litchi_ole_common::ole_streams::{
    CF_DIB, ClipboardFormat, Limits as OleStreamLimits, OleNativeStream, OlePresentationStream,
};
use std::io::Cursor;

#[test]
fn object_pool_target_names_follow_decimal_storage_form() {
    assert!(is_object_storage_name("_0"));
    assert!(is_object_storage_name("_00042"));
    assert!(is_object_storage_name("_-1"));
    assert!(!is_object_storage_name("Object"));
    assert!(!is_object_storage_name("_"));
    assert!(!is_object_storage_name("_+1"));
    assert!(!is_object_storage_name("_42x"));
}

#[test]
fn target_discovery_keeps_exact_object_pool_storage_names() {
    let mut writer = OleWriter::new();
    writer.create_storage(&[OBJECT_POOL, "_00042"]).unwrap();
    writer.create_storage(&[OBJECT_POOL, "_-1"]).unwrap();
    writer.create_storage(&[OBJECT_POOL, "not-an-id"]).unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let (targets, object_pool_exists) = discover_targets(&bytes.into_inner(), Limits::default())
        .expect("ObjectPool target discovery should succeed");
    assert!(object_pool_exists);
    assert_eq!(targets.len(), 2);
    assert!(
        targets.get("_00042").is_some_and(|target| {
            target.path() == [OBJECT_POOL.to_owned(), "_00042".to_owned()]
        })
    );
    assert!(
        targets
            .get("_-1")
            .is_some_and(|target| { target.path() == [OBJECT_POOL.to_owned(), "_-1".to_owned()] })
    );
}

#[test]
fn obj_info_reads_the_doc_opaque_stream_shape() {
    let info = Info::read(&[0x00, 0x82, 0x03, 0x00, 0x00, 0x00]).unwrap();
    assert!(info.recompose_on_resize);
    assert!(info.view_object);
    assert_eq!(info.clipboard_format, 3);
    assert!(info.persist2_present);
    assert_eq!(
        info.to_bytes().unwrap(),
        [0x00, 0x82, 0x03, 0x00, 0x00, 0x00]
    );
    assert!(Info::read(&[0x00, 0x04, 0x00, 0x00]).is_err());
}

#[test]
fn obj_info_preserves_undefined_bits_and_optional_presence() {
    let bytes = [0x2D, 0x40, 0x14, 0x00, 0xF0, 0x00];
    let info = Info::read(&bytes).unwrap();
    assert_eq!(info.reserved_persist1, 0x402D);
    assert_eq!(info.reserved_persist2, 0x00F0);
    assert!(info.persist2_present);
    assert_eq!(info.to_bytes().unwrap(), bytes);

    let without_optional = Info::read(&[0x00, 0x00, 0x03, 0x00]).unwrap();
    assert!(!without_optional.persist2_present);
    assert_eq!(
        without_optional.to_bytes().unwrap(),
        [0x00, 0x00, 0x03, 0x00]
    );

    let explicit_zero_optional = Info::read(&[0x00, 0x00, 0x03, 0x00, 0x00, 0x00]).unwrap();
    assert!(explicit_zero_optional.persist2_present);
    assert_eq!(
        explicit_zero_optional.to_bytes().unwrap(),
        [0x00, 0x00, 0x03, 0x00, 0x00, 0x00]
    );
}

#[test]
fn obj_info_rejects_invalid_required_bits_without_ole_access() {
    assert!(Info::read(&[0x00, 0x08, 0x00, 0x00]).is_err());
    assert!(Info::read(&[0x00, 0x00, 0x00, 0x00, 0x02, 0x00]).is_err());
    assert!(Info::read(&[0x00, 0x20, 0x00, 0x00]).is_err());

    let mut info = Info::read(&[0x00, 0x00, 0x00, 0x00]).unwrap();
    info.reserved_persist1 = 1 << 1;
    assert!(info.to_bytes().is_err());
    info.reserved_persist1 = 0;
    info.reserved_persist2 = 1 << 1;
    assert!(info.to_bytes().is_err());
}

#[test]
fn field_validation_rejects_orphan_and_unclosed_markers() {
    assert!(
        validate_existing_fields(
            &[FieldMarker {
                cp: 0,
                descriptor: [0x14, 0],
            }],
            1,
        )
        .is_err()
    );
    assert!(
        validate_existing_fields(
            &[FieldMarker {
                cp: 0,
                descriptor: [0x13, 0x3A],
            }],
            1,
        )
        .is_err()
    );
}

#[test]
fn snapshot_clones_share_the_first_public_source_view_allocation() {
    let source = Snapshot::open(base_doc(), Limits::default()).unwrap();
    let bytes_clone = source.clone();
    let shared_clone = source.clone();

    let bytes = bytes_clone.bytes();
    let shared = shared_clone.bytes_shared();
    assert_eq!(bytes, shared.as_ref());
    assert_eq!(bytes.as_ptr(), shared.as_ptr());

    let bytes_clone_shared = bytes_clone.bytes_shared();
    assert!(std::sync::Arc::ptr_eq(&bytes_clone_shared, &shared));
    assert_eq!(shared_clone.bytes().as_ptr(), shared.as_ptr());
}

fn base_doc() -> Vec<u8> {
    let mut writer = Writer::new();
    writer.add_paragraph("embedded metadata").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    normalize_word2002_dop(output.into_inner())
}

fn normalize_word2002_dop(bytes: Vec<u8>) -> Vec<u8> {
    let mut package = ObjectEditor::open(bytes, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let word = package.stream(&word_path).unwrap();
    let fib = crate::parts::fib::FileInformationBlock::parse(word).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut word = word.to_vec();
    let mut table = package.stream(&table_path).unwrap().to_vec();
    let (offset, length) = fib.get_table_pointer(31).unwrap();
    let offset = usize::try_from(offset).unwrap();
    let length = usize::try_from(length).unwrap();
    let dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
        false, false, false, true,
    );
    if length < dop.len() {
        let insertion = offset + length;
        let extra = dop.len() - length;
        table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
        let count = fib.table_pointer_count().unwrap();
        for index in 0..count {
            let pointer = 154 + index * 8;
            let current = usize::try_from(u32::from_le_bytes(
                word[pointer..pointer + 4].try_into().unwrap(),
            ))
            .unwrap();
            if current >= insertion {
                let shifted = u32::try_from(current + extra).unwrap();
                word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
            }
        }
    }
    table[offset..offset + dop.len()].copy_from_slice(&dop);
    let pointer = 154 + 31 * 8;
    word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
    package.put_stream(&word_path, word).unwrap();
    package.put_stream(&table_path, table).unwrap();
    package.finish().unwrap()
}

fn protected_doc(source: Vec<u8>) -> Vec<u8> {
    const DOP_INDEX: usize = 31;
    let mut package = ObjectEditor::open(source, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let table_name = {
        let word = package.stream(&word_path).unwrap();
        let fib = crate::parts::fib::FileInformationBlock::parse(word).unwrap();
        if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        }
        .to_owned()
    };
    let table_path = [table_name];
    let mut word = package.stream(&word_path).unwrap().to_vec();
    let mut table = package.stream(&table_path).unwrap().to_vec();
    let offset = u32::try_from(table.len()).unwrap();
    let mut dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
        false, false, false, true,
    );
    dop[6] = 0x10;
    table.extend_from_slice(&dop);
    let pointer = 154 + DOP_INDEX * 8;
    word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
    word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
    package.put_stream(&word_path, word).unwrap();
    package.put_stream(&table_path, table).unwrap();
    package.finish().unwrap()
}

fn picture_data(marker: u32) -> Vec<u8> {
    let mut value = 12u32.to_le_bytes().to_vec();
    value.extend_from_slice(&marker.to_le_bytes());
    value.extend_from_slice(&[0; 4]);
    value
}

fn object_cfb(comp_obj: &[u8], ole: &[u8], obj_info: &[u8], unknown: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["\u{1}CompObj"], comp_obj).unwrap();
    writer.create_stream(&["\u{1}Ole"], ole).unwrap();
    writer.create_stream(&["\u{3}ObjInfo"], obj_info).unwrap();
    writer.create_stream(&["VendorMetadata"], unknown).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn object_cfb_with_payload(
    comp_obj: &[u8],
    ole: &[u8],
    obj_info: &[u8],
    unknown: &[u8],
    payload: &[u8],
) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["\u{1}CompObj"], comp_obj).unwrap();
    writer.create_stream(&["\u{1}Ole"], ole).unwrap();
    writer.create_stream(&["\u{3}ObjInfo"], obj_info).unwrap();
    writer.create_stream(&["VendorMetadata"], unknown).unwrap();
    writer.create_storage(&["OpaquePayload"]).unwrap();
    writer
        .create_stream(&["OpaquePayload", "Binary"], payload)
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn object_cfb_with_ole2_streams(
    comp_obj: &[u8],
    ole: &[u8],
    obj_info: &[u8],
    unknown: &[u8],
    presentation: &[u8],
    native: &[u8],
) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["\u{1}CompObj"], comp_obj).unwrap();
    writer.create_stream(&["\u{1}Ole"], ole).unwrap();
    writer.create_stream(&["\u{3}ObjInfo"], obj_info).unwrap();
    writer.create_stream(&["VendorMetadata"], unknown).unwrap();
    writer
        .create_stream(&["\u{2}OlePres000"], presentation)
        .unwrap();
    writer.create_stream(&["\u{1}Ole10Native"], native).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ole2_test_object(marker: u8) -> Vec<u8> {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00, 0x00, 0x00];
    let presentation = OlePresentationStream::new(
        ClipboardFormat::Standard(CF_DIB),
        [marker, marker.wrapping_add(1)],
    )
    .unwrap()
    .to_bytes()
    .unwrap();
    let native = OleNativeStream::new([marker.wrapping_add(2), marker.wrapping_add(3)])
        .unwrap()
        .to_bytes()
        .unwrap();
    object_cfb_with_ole2_streams(
        &comp_obj,
        &ole,
        &obj_info,
        &[marker, marker.wrapping_add(4)],
        &presentation,
        &native,
    )
}

#[test]
fn inventory_exposes_inert_ole_metadata_and_unknown_streams() {
    let mut comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    comp_obj.extend_from_slice(&[0xA1, 0xB2, 0xC3]);
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x92, 0x03, 0x00, 0x00, 0x00];
    let unknown = [0x10, 0x20, 0x30, 0x40];
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            77,
            object_cfb(&comp_obj, &ole, &obj_info, &unknown),
            picture_data(77),
        ))
        .unwrap();

    let inventory = editor.inventory().unwrap();
    let entry = inventory.get(77).expect("metadata entry");
    let metadata = entry.metadata();
    let comp_obj = metadata.comp_obj().expect("CompObj metadata");
    assert_eq!(comp_obj.ansi_user_type(), "Microsoft Word Document");
    assert!(comp_obj.has_reserved_ansi());
    assert!(comp_obj.has_reserved_unicode());
    assert!(comp_obj.bytes().ends_with(&[0xA1, 0xB2, 0xC3]));
    assert_eq!(comp_obj.trailing(), &[0xA1, 0xB2, 0xC3]);
    let ole_metadata = metadata.ole().expect("Ole metadata");
    assert_eq!(ole_metadata.kind(), Kind::Embedded);
    assert_eq!(ole_metadata.bytes(), ole);
    assert!(metadata.is_activex());
    assert_eq!(
        metadata
            .obj_info()
            .expect("ObjInfo metadata")
            .clipboard_format,
        3
    );
    assert_eq!(metadata.unknown().len(), 1);
    assert_eq!(metadata.unknown()[0].path(), &["VendorMetadata".to_owned()]);
    assert_eq!(metadata.unknown()[0].bytes(), unknown);
    assert!(metadata.has_unknown());

    let snapshot = inventory.clone();
    let invalid = super::WriteOptions::new(78, vec![0], picture_data(78));
    assert!(editor.add(invalid).is_err());
    assert_eq!(editor.inventory().unwrap(), snapshot);

    let bytes = editor.finish().unwrap();
    let reopened = super::Editor::open(bytes, Limits::default()).unwrap();
    assert_eq!(reopened.inventory().unwrap(), snapshot);
}

#[test]
fn malformed_known_metadata_remains_lossless_unknown_bytes() {
    let malformed = [0xEE, 0xDD, 0xCC, 0xBB, 0xAA];
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x82, 0x03, 0x00, 0x00, 0x00];
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            88,
            object_cfb(&malformed, &ole, &obj_info, &[]),
            picture_data(88),
        ))
        .unwrap();

    let inventory = editor.inventory().unwrap();
    let entry = inventory.get(88).expect("metadata entry");
    assert!(entry.metadata().comp_obj().is_none());
    let unknown = entry
        .metadata()
        .unknown()
        .iter()
        .find(|value| value.name() == Some("\u{1}CompObj"))
        .expect("malformed CompObj bytes");
    assert_eq!(unknown.bytes(), malformed);
}

#[test]
fn snapshot_transactions_are_source_checked_and_preserve_opaque_object_data() {
    let mut comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    comp_obj.extend_from_slice(&[0xA1, 0xB2, 0xC3]);
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00, 0x00, 0x00];
    let unknown = [0x10, 0x20, 0x30, 0x40];
    let payload = [0xDE, 0xAD, 0xBE, 0xEF, 0x00];

    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            77,
            object_cfb_with_payload(&comp_obj, &ole, &obj_info, &unknown, &payload),
            picture_data(77),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), Limits::default()).unwrap();

    let mut transaction = source.edit();
    transaction
        .update_link(77, |link| {
            link.set_cache_hint(true);
            Ok(())
        })
        .unwrap();
    transaction
        .update_info(77, |info| info.display_as_icon = true)
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_noop());
    assert!(
        !source
            .inventory()
            .unwrap()
            .get(77)
            .unwrap()
            .metadata()
            .obj_info()
            .unwrap()
            .display_as_icon
    );
    assert!(
        commit
            .snapshot()
            .inventory()
            .unwrap()
            .get(77)
            .unwrap()
            .metadata()
            .obj_info()
            .unwrap()
            .display_as_icon
    );
    assert_eq!(
        commit
            .snapshot()
            .inventory()
            .unwrap()
            .get(77)
            .unwrap()
            .metadata()
            .ole()
            .unwrap()
            .flags()
            & 0x1000,
        0x1000
    );

    let mut ole_file = litchi_cfb::OleFile::open(Cursor::new(commit.snapshot().bytes())).unwrap();
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_77", "VendorMetadata"])
            .unwrap(),
        unknown
    );
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_77", "OpaquePayload", "Binary"])
            .unwrap(),
        payload
    );

    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(&applied, commit.snapshot());
    let reverted = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(reverted, source);

    let stale = Snapshot::open(base_doc(), Limits::default()).unwrap();
    assert!(matches!(
        commit.patch().apply(&stale),
        Err(TransactionError::Conflict)
    ));
}

#[test]
fn doc_ole2_stream_access_edits_atomically_and_preserves_host_metadata() {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00, 0x00, 0x00];
    let unknown = [0xFA, 0xCE, 0xD0, 0x0C];
    let presentation =
        OlePresentationStream::new(ClipboardFormat::Standard(CF_DIB), [0x10, 0x20, 0x30])
            .unwrap()
            .to_bytes()
            .unwrap();
    let native = OleNativeStream::new([0x40, 0x50, 0x60])
        .unwrap()
        .to_bytes()
        .unwrap();

    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            77,
            object_cfb_with_ole2_streams(
                &comp_obj,
                &ole,
                &obj_info,
                &unknown,
                &presentation,
                &native,
            ),
            picture_data(77),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), Limits::default()).unwrap();
    let reference = source.objects().unwrap().into_iter().next().unwrap();

    let source_presentation = source.presentation(77, 0).unwrap().unwrap();
    assert_eq!(source_presentation.data(), [0x10, 0x20, 0x30]);
    assert_eq!(
        source
            .presentation_for(&reference, 0)
            .unwrap()
            .unwrap()
            .bytes(),
        source_presentation.bytes()
    );
    assert_eq!(
        source.native(77).unwrap().unwrap().data(),
        [0x40, 0x50, 0x60]
    );
    assert_eq!(
        source.native_for(&reference).unwrap().unwrap().data(),
        [0x40, 0x50, 0x60]
    );

    let small = OleStreamLimits {
        max_bytes: presentation.len(),
        max_data_bytes: 2,
        max_toc_entries: 1,
    };
    assert!(source.presentation_with_limits(77, 0, small).is_err());
    assert!(source.native_with_limits(77, small).is_err());

    let mut stream_edit = source_presentation.edit();
    stream_edit.set_width(640).unwrap();
    stream_edit.set_height(480).unwrap();
    stream_edit.set_data([0xAA, 0xBB]).unwrap();
    let stream_commit = stream_edit.commit().unwrap();

    let mut transaction = source.edit();
    transaction
        .apply_presentation_patch_for(&reference, 0, stream_commit.patch())
        .unwrap();
    transaction
        .update_native_for(&reference, |native| {
            native.set_data([0xCC, 0xDD]).map(|_| ())
        })
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_noop());
    assert_eq!(
        source.presentation(77, 0).unwrap().unwrap().data(),
        [0x10, 0x20, 0x30]
    );
    let after = commit.snapshot();
    let after_presentation = after.presentation(77, 0).unwrap().unwrap();
    assert_eq!(after_presentation.data(), [0xAA, 0xBB]);
    assert_eq!(after_presentation.width(), 640);
    assert_eq!(after_presentation.height(), 480);
    assert_eq!(after.native(77).unwrap().unwrap().data(), [0xCC, 0xDD]);

    let mut ole_file = litchi_cfb::OleFile::open(Cursor::new(after.bytes())).unwrap();
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_77", "\u{1}CompObj"])
            .unwrap(),
        comp_obj
    );
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_77", "VendorMetadata"])
            .unwrap(),
        unknown
    );

    let reverted = commit.patch().inverse().apply(after).unwrap();
    assert_eq!(reverted, source);
    let stale = Snapshot::open(base_doc(), Limits::default()).unwrap();
    assert!(matches!(
        commit.patch().apply(&stale),
        Err(TransactionError::Conflict)
    ));

    let mut typed_stale = after.edit();
    assert!(
        typed_stale
            .apply_presentation_patch(77, 0, stream_commit.patch())
            .is_err()
    );
    assert!(!typed_stale.is_changed().unwrap());

    let mut noop = source.edit();
    noop.update_presentation(77, 0, |_| Ok(())).unwrap();
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert!(noop_commit.patch().is_noop());
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &noop_commit.snapshot().bytes_shared()
    ));

    let original_presentation_data = source_presentation.data().to_vec();
    let original_presentation_width = source_presentation.width();
    let original_presentation_height = source_presentation.height();
    let original_native_data = source.native(77).unwrap().unwrap().data().to_vec();
    let mut reverted = source.edit();
    reverted
        .update_presentation(77, 0, |presentation| {
            presentation.set_width(640)?;
            presentation.set_height(480)?;
            presentation.set_data([0x01, 0x02])?;
            presentation.set_width(original_presentation_width)?;
            presentation.set_height(original_presentation_height)?;
            presentation.set_data(original_presentation_data)?;
            Ok(())
        })
        .unwrap();
    reverted
        .update_native(77, |native| {
            native.set_data([0x03, 0x04])?;
            native.set_data(original_native_data)?;
            Ok(())
        })
        .unwrap();
    let reverted_commit = reverted.commit().unwrap();
    assert!(!reverted_commit.changed());
    assert!(reverted_commit.patch().is_noop());
    assert_eq!(reverted_commit.snapshot().bytes(), source.bytes());
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &reverted_commit.snapshot().bytes_shared()
    ));

    let mut malformed_editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    malformed_editor
        .add(super::WriteOptions::new(
            88,
            object_cfb_with_ole2_streams(&comp_obj, &ole, &obj_info, &unknown, &[0x00], &native),
            picture_data(88),
        ))
        .unwrap();
    let malformed = Snapshot::open(malformed_editor.finish().unwrap(), Limits::default()).unwrap();
    assert!(malformed.presentation(88, 0).is_err());
}

#[test]
fn embedded_patch_cannot_cross_into_an_enforcing_destination() {
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            77,
            ole2_test_object(0x40),
            picture_data(77),
        ))
        .unwrap();
    let protected = protected_doc(editor.finish().unwrap());
    let authorization = crate::parts::protection::ProtectionAuthorization::audited(
        "alice",
        "approved embedded metadata patch",
    )
    .unwrap();
    let allowed = Snapshot::open_with_policy(
        protected.clone(),
        Limits::default(),
        crate::parts::protection::ProtectionPolicy::allow_protected(authorization),
    )
    .unwrap();
    let enforcing = Snapshot::open(protected, Limits::default()).unwrap();

    let mut transaction = allowed.edit();
    transaction
        .update_info(77, |info| {
            info.display_as_icon = true;
        })
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(matches!(
        commit.patch().apply(&enforcing),
        Err(TransactionError::Invalid(
            crate::package::Error::ProtectionDenied(_)
        ))
    ));

    let changed_enforcing =
        Snapshot::open(commit.snapshot().bytes().to_vec(), Limits::default()).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&changed_enforcing),
        Err(TransactionError::Invalid(
            crate::package::Error::ProtectionDenied(_)
        ))
    ));
}

#[test]
fn reference_selected_streams_reject_forged_and_ambiguous_targets_atomically() {
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            77,
            ole2_test_object(0x10),
            picture_data(77),
        ))
        .unwrap();
    editor
        .add(super::WriteOptions::new(
            88,
            ole2_test_object(0x20),
            picture_data(88),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), Limits::default()).unwrap();
    let references = source.objects().unwrap();
    let first = references
        .iter()
        .find(|reference| reference.storage_id == 77)
        .cloned()
        .unwrap();
    let second = references
        .iter()
        .find(|reference| reference.storage_id == 88)
        .cloned()
        .unwrap();

    let mut forged_name = first.clone();
    forged_name.storage_name = second.storage_name.clone();
    assert!(source.native_for(&forged_name).is_err());
    let mut transaction = source.edit();
    assert!(
        transaction
            .update_native_for(&forged_name, |native| {
                native.set_data([0xF0, 0xF1]).map(|_| ())
            })
            .is_err()
    );
    assert!(!transaction.is_changed().unwrap());
    assert_eq!(transaction.snapshot().unwrap(), source);

    let mut forged_id = first.clone();
    forged_id.storage_id = second.storage_id;
    assert!(source.presentation_for(&forged_id, 0).is_err());
    assert_eq!(source.native(77).unwrap().unwrap().data(), [0x12, 0x13]);
    assert_eq!(source.native(88).unwrap().unwrap().data(), [0x22, 0x23]);

    let target = Target::new("_77", [OBJECT_POOL, "_77"]).unwrap();
    let mut object_editor = ObjectEditor::open(
        source.bytes().to_vec(),
        Targets::one(target),
        Limits::default(),
    )
    .unwrap();
    object_editor
        .add_storage(
            Target::new("_00077", [OBJECT_POOL, "_00077"]).unwrap(),
            ole2_test_object(0x30),
        )
        .unwrap();
    let ambiguous = Snapshot::open(object_editor.finish().unwrap(), Limits::default()).unwrap();
    let ole_file = litchi_cfb::OleFile::open(Cursor::new(ambiguous.bytes())).unwrap();
    assert!(
        ole_file
            .list_directory_entries(&[OBJECT_POOL])
            .unwrap()
            .iter()
            .any(|entry| entry.name == "_00077")
    );

    let ambiguous_reference = ambiguous
        .objects()
        .unwrap()
        .into_iter()
        .find(|reference| reference.storage_id == 77)
        .unwrap();
    assert!(ambiguous.native_for(&ambiguous_reference).is_err());
    let mut ambiguous_transaction = ambiguous.edit();
    assert!(
        ambiguous_transaction
            .update_native_for(&ambiguous_reference, |native| {
                native.set_data([0xF2, 0xF3]).map(|_| ())
            })
            .is_err()
    );
    assert!(!ambiguous_transaction.is_changed().unwrap());
    assert_eq!(ambiguous_transaction.snapshot().unwrap(), ambiguous);
}

#[test]
fn snapshot_transactions_keep_invalid_metadata_and_storage_edits_atomic() {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00, 0x00, 0x00];
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            91,
            object_cfb(&comp_obj, &ole, &obj_info, &[0xAB, 0xCD]),
            picture_data(91),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), Limits::default()).unwrap();

    let mut transaction = source.edit();
    assert!(
        transaction
            .update_info(91, |info| {
                info.stream_control = true;
                info.activex = false;
            })
            .is_err()
    );
    assert!(!transaction.is_changed().unwrap());
    assert_eq!(transaction.snapshot().unwrap(), source);

    assert!(transaction.replace_storage(91, vec![0x00]).is_err());
    assert!(!transaction.is_changed().unwrap());
    assert_eq!(transaction.snapshot().unwrap(), source);

    let mut noop_transaction = source.edit();
    noop_transaction.update_info(91, |_| {}).unwrap();
    let commit = noop_transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().apply(&source).unwrap(), source);
}

#[test]
fn storage_replacement_admission_rejects_malformed_and_oversized_before_publish() {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00];
    let mut limits = Limits::default();
    limits.max_object_size = 16 * 1024;
    limits.max_streams_per_object = 4;
    let mut editor = super::Editor::open(base_doc(), limits).unwrap();
    editor
        .add(super::WriteOptions::new(
            92,
            object_cfb(&comp_obj, &ole, &obj_info, &[0xAB, 0xCD]),
            picture_data(92),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), limits).unwrap();

    let mut malformed = source.edit();
    assert!(malformed.replace_storage(92, vec![0x00, 0x01]).is_err());
    assert!(!malformed.is_changed().unwrap());
    let malformed_snapshot = malformed.snapshot().unwrap();
    assert_eq!(malformed_snapshot, source);
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &malformed_snapshot.bytes_shared()
    ));

    let oversized = object_cfb_with_payload(
        &comp_obj,
        &ole,
        &obj_info,
        &[0xF0],
        &vec![0u8; usize::try_from(limits.max_object_size).unwrap()],
    );
    let mut oversized_edit = source.edit();
    assert!(oversized_edit.replace_storage(92, oversized).is_err());
    assert!(!oversized_edit.is_changed().unwrap());
    let oversized_snapshot = oversized_edit.snapshot().unwrap();
    assert_eq!(oversized_snapshot, source);
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &oversized_snapshot.bytes_shared()
    ));

    let too_many_streams = object_cfb_with_payload(&comp_obj, &ole, &obj_info, &[0xF1], &[0xF2]);
    let mut per_object_edit = source.edit();
    assert!(
        per_object_edit
            .replace_storage(92, too_many_streams)
            .is_err()
    );
    assert!(!per_object_edit.is_changed().unwrap());
    let per_object_snapshot = per_object_edit.snapshot().unwrap();
    assert_eq!(per_object_snapshot, source);
    assert!(std::sync::Arc::ptr_eq(
        &source.bytes_shared(),
        &per_object_snapshot.bytes_shared()
    ));
}

#[test]
fn replacing_a_storage_keeps_the_field_reference_and_reparses_opaque_payloads() {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00];
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            101,
            object_cfb_with_payload(&comp_obj, &ole, &obj_info, &[0x01], &[0x02]),
            picture_data(101),
        ))
        .unwrap();
    let source = Snapshot::open(editor.finish().unwrap(), Limits::default()).unwrap();
    let replacement = object_cfb_with_payload(&comp_obj, &ole, &obj_info, &[0xF0], &[0xF1, 0xF2]);

    let mut transaction = source.edit();
    transaction.replace_storage(101, replacement).unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.snapshot().objects().unwrap()[0].storage_id, 101);

    let mut ole_file = litchi_cfb::OleFile::open(Cursor::new(commit.snapshot().bytes())).unwrap();
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_101", "VendorMetadata"])
            .unwrap(),
        [0xF0]
    );
    assert_eq!(
        ole_file
            .open_stream(&["ObjectPool", "_101", "OpaquePayload", "Binary"])
            .unwrap(),
        [0xF1, 0xF2]
    );
}

#[test]
fn snapshot_rejects_a_field_whose_objectpool_owner_was_removed() {
    let comp_obj = crate::writer::ole_metadata::generate_compobj_stream();
    let ole = crate::writer::ole_metadata::generate_ole_stream();
    let obj_info = [0x00, 0x00, 0x03, 0x00];
    let mut editor = super::Editor::open(base_doc(), Limits::default()).unwrap();
    editor
        .add(super::WriteOptions::new(
            123,
            object_cfb(&comp_obj, &ole, &obj_info, &[]),
            picture_data(123),
        ))
        .unwrap();
    let bytes = editor.finish().unwrap();

    let target = Target::new("_123", [OBJECT_POOL, "_123"]).unwrap();
    let mut object_editor =
        ObjectEditor::open(bytes, Targets::one(target), Limits::default()).unwrap();
    object_editor.remove_storage("_123").unwrap();
    let orphaned = object_editor.finish().unwrap();

    assert!(Snapshot::open(orphaned, Limits::default()).is_err());
}
