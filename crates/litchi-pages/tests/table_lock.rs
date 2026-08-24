use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, Limits as ArchiveLimits, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{tp, tsa, tsd, tsp, tst, tswp};
use litchi_pages::table::lock::State;
use litchi_pages::{
    BodyTableLockError, BodyTableLockLimitKind, BodyTableSelector, Limits, Package,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const BODY_IDENTIFIER: u64 = 42;
const ATTACHMENT_IDENTIFIER: u64 = 100;
const DRAWABLE_IDENTIFIER: u64 = 200;
const MODEL_IDENTIFIER: u64 = 300;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const LEGACY_TABLE_INFO_MESSAGE_TYPE: u32 = 6_003;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const UNKNOWN_TABLE_INFO_FIELD: u32 = 99;
const ROOT_BODY_FIELD: u32 = 4;
const DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const TABLE_INFO_SUPER_FIELD: u32 = 1;
const TABLE_INFO_MODEL_FIELD: u32 = 2;
const TABLE_BODY_FIELD: u32 = 9;

type TestResult<T> = Result<T, Box<dyn StdError>>;

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn field_reference(path: impl Into<FieldPath>, identifier: u64) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.push(identifier);
    field
}

fn synthetic_package(locked: Option<bool>) -> TestResult<Vec<u8>> {
    synthetic_package_with_table_count(locked, 1)
}

fn synthetic_package_with_table_count(
    locked: Option<bool>,
    table_count: usize,
) -> TestResult<Vec<u8>> {
    synthetic_package_with_table_count_mode(locked, table_count, false)
}

fn synthetic_package_with_unique_table_count(
    locked: Option<bool>,
    table_count: usize,
) -> TestResult<Vec<u8>> {
    synthetic_package_with_table_count_mode(locked, table_count, true)
}

fn synthetic_package_with_table_count_mode(
    locked: Option<bool>,
    table_count: usize,
    unique_table_ids: bool,
) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };

    let mut body_text = String::from("A");
    for _ in 0..table_count {
        body_text.push('\u{fffc}');
    }
    let table_ids = |index: usize| {
        let offset = if unique_table_ids {
            u64::try_from(index).unwrap_or(u64::MAX).saturating_mul(3)
        } else {
            0
        };
        (
            ATTACHMENT_IDENTIFIER.saturating_add(offset),
            DRAWABLE_IDENTIFIER.saturating_add(offset),
            MODEL_IDENTIFIER.saturating_add(offset),
        )
    };
    let attachment_identifiers = (0..table_count)
        .map(|index| table_ids(index).0)
        .collect::<Vec<_>>();
    let table_entries = (0..table_count)
        .map(|index| {
            let attachment_identifier = table_ids(index).0;
            tswp::object_attribute_table::ObjectAttribute {
                character_index: u32::try_from(index + 1).unwrap_or(u32::MAX),
                object: Some(reference(attachment_identifier)),
            }
        })
        .collect();
    let body_payload = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec![body_text],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: table_entries,
        }),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();

    let mut root_object = object(
        1,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));
    let mut body_object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body_payload,
        &attachment_identifiers,
    )?;
    for attachment_identifier in &attachment_identifiers {
        body_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(
                vec![TABLE_BODY_FIELD],
                *attachment_identifier,
            ));
    }
    let mut objects = vec![root_object, body_object];
    let object_count = if unique_table_ids {
        table_count
    } else {
        table_count.min(1)
    };
    for index in 0..object_count {
        let (attachment_identifier, drawable_identifier, model_identifier) = table_ids(index);
        let attachment_payload = tswp::DrawableAttachmentArchive {
            drawable: Some(reference(drawable_identifier)),
            ..tswp::DrawableAttachmentArchive::default()
        }
        .encode_to_vec();
        let mut table_info_payload = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(BODY_IDENTIFIER)),
                locked,
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(model_identifier),
            ..tst::TableInfoArchive::default()
        }
        .encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(
            &mut table_info_payload,
            UNKNOWN_TABLE_INFO_FIELD,
            0xfeed_beef,
        )?;
        let table_name = if index == 0 {
            "Revenue".to_owned()
        } else {
            format!("Revenue {index}")
        };
        let model_payload = tst::TableModelArchive {
            table_name,
            ..tst::TableModelArchive::default()
        }
        .encode_to_vec();
        let mut attachment_object = object(
            attachment_identifier,
            ATTACHMENT_MESSAGE_TYPE,
            attachment_payload,
            &[drawable_identifier],
        )?;
        attachment_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![1], drawable_identifier));
        let mut drawable_object = object(
            drawable_identifier,
            TABLE_INFO_MESSAGE_TYPE,
            table_info_payload,
            &[BODY_IDENTIFIER, model_identifier],
        )?;
        drawable_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![1, 2], BODY_IDENTIFIER),
                field_reference(vec![2], model_identifier),
            ]);
        objects.push(attachment_object);
        objects.push(drawable_object);
        objects.push(object(
            model_identifier,
            TABLE_MODEL_MESSAGE_TYPE,
            model_payload,
            &[],
        )?);
    }
    let archive = Archive { objects };
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"untouched".as_slice()),
            (DOCUMENT_MEMBER, component.as_slice()),
        ],
        Limits::default(),
    )?)
}

#[test]
fn oversized_body_table_entry_list_is_rejected_before_allocation() -> TestResult<()> {
    let source = synthetic_package_with_table_count(None, 4_097)?;
    let package = Package::from_bytes(&source)?;

    assert!(matches!(
        package.body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::LimitExceeded {
            kind: BodyTableLockLimitKind::PayloadItems,
            observed: 4_097,
            maximum: 4_096,
        })
    ));
    Ok(())
}

#[test]
fn cumulative_table_info_wire_budget_has_a_bounded_boundary() -> TestResult<()> {
    let source = synthetic_package_with_unique_table_count(None, 8)?;
    let mut wire_field_boundary = None;
    let mut successful_fields = None;

    for fields in 1..=4_096 {
        let archive_limits = ArchiveLimits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        let Ok(package) = Package::from_bytes_with_limits(&source, limits) else {
            continue;
        };
        match package.body_table_lock(BodyTableSelector::index(0)) {
            Ok(State::Unlocked | State::Locked) => {
                successful_fields = Some(fields);
                break;
            },
            Err(BodyTableLockError::LimitExceeded {
                kind: BodyTableLockLimitKind::WireFields,
                ..
            }) if wire_field_boundary.is_none() => wire_field_boundary = Some(fields),
            Err(_) => {},
        }
    }

    assert!(successful_fields.is_some());
    assert!(wire_field_boundary.is_some());
    assert!(wire_field_boundary < successful_fields);
    Ok(())
}

#[test]
fn cumulative_table_info_wire_budget_scales_with_unique_tables() -> TestResult<()> {
    let one_source = synthetic_package_with_unique_table_count(None, 1)?;
    let two_source = synthetic_package_with_unique_table_count(None, 2)?;
    let mut one_success = None;
    let mut two_success = None;

    for fields in 1..=4_096 {
        let archive_limits = ArchiveLimits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        if one_success.is_none()
            && let Ok(one) = Package::from_bytes_with_limits(&one_source, limits)
            && matches!(
                one.body_table_lock(BodyTableSelector::index(0)),
                Ok(State::Unlocked) | Ok(State::Locked)
            )
        {
            one_success = Some(fields);
        }
        if two_success.is_none()
            && let Ok(two) = Package::from_bytes_with_limits(&two_source, limits)
            && matches!(
                two.body_table_lock(BodyTableSelector::index(0)),
                Ok(State::Unlocked) | Ok(State::Locked)
            )
        {
            two_success = Some(fields);
        }
        if one_success.is_some() && two_success.is_some() {
            break;
        }
    }

    assert!(one_success.is_some());
    assert!(two_success.is_some());
    assert!(one_success < two_success);
    Ok(())
}

fn table_info_payload(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object(DRAWABLE_IDENTIFIER)
        .ok_or("missing drawable")?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or("missing table info")?
        .data
        .clone())
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    mutate(&mut archive)?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn noncanonical_document_prefix(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut bytes = stream.as_bytes().to_vec();
    let (_, prefix_length) = decode_varint_from_bytes(&bytes)?;
    let mut prefix = bytes
        .get(..prefix_length)
        .ok_or("missing document prefix")?
        .to_vec();
    let last = prefix.last_mut().ok_or("empty document prefix")?;
    *last |= 0x80;
    prefix.push(0);
    bytes.splice(0..prefix_length, prefix);
    let component = SnappyStream::compress(&bytes)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn rewrite_message_references(
    archive: &mut Archive,
    object_identifier: u64,
    message_type: u32,
    references: &[u64],
) -> TestResult<()> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or("missing metadata object")?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or("missing metadata message")?;
    object.archive_info.message_infos[message_index].object_references = references.to_vec();
    Ok(())
}

fn rewrite_message_metadata(
    archive: &mut Archive,
    object_identifier: u64,
    message_type: u32,
    mutate: impl FnOnce(&mut litchi_iwa_core::MessageInfo),
) -> TestResult<()> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or("missing metadata object")?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or("missing metadata message")?;
    mutate(&mut object.archive_info.message_infos[message_index]);
    Ok(())
}

fn assert_invalid_metadata_source(source: &[u8]) -> TestResult<()> {
    let package = Package::from_bytes(source)?;
    assert!(matches!(
        package.body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    Ok(())
}

fn lock_field(payload: &[u8]) -> TestResult<Option<bool>> {
    let outer = WireView::parse(payload)?;
    let super_field = outer
        .fields()
        .find(|field| field.number() == 1)
        .ok_or("missing table-info super")?;
    let drawable = WireView::parse(super_field.payload())?;
    let Some(field) = drawable.fields().find(|field| field.number() == 5) else {
        return Ok(None);
    };
    let (value, width) = decode_varint_from_bytes(field.payload())?;
    if width != field.payload().len() {
        return Err("non-canonical lock field".into());
    }
    Ok(Some(match value {
        0 => false,
        1 => true,
        _ => return Err("non-boolean lock field".into()),
    }))
}

fn has_unknown_table_info_field(payload: &[u8]) -> TestResult<bool> {
    Ok(WireView::parse(payload)?.fields().any(|field| {
        field.number() == UNKNOWN_TABLE_INFO_FIELD
            && decode_varint_from_bytes(field.payload())
                .is_ok_and(|(value, width)| value == 0xfeed_beef && width == field.payload().len())
    }))
}

fn sentinel_payload(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == "Data/sentinel.bin")
        .ok_or("missing sentinel member")?
        .data()
        .to_vec())
}

#[test]
fn absent_and_explicit_false_are_unlocked_noop_states() -> TestResult<()> {
    for locked in [None, Some(false)] {
        let source = synthetic_package(locked)?;
        let package = Package::from_bytes(&source)?;
        assert_eq!(
            package.body_table_lock(BodyTableSelector::name("Revenue"))?,
            State::Unlocked
        );

        let mut edit = package.edit_body_table_lock(BodyTableSelector::index(0))?;
        edit.unlock();
        let commit = edit.commit()?;
        assert_eq!(commit.package().exact_bytes(), source.as_slice());
        assert!(commit.patch().is_noop());
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(
            lock_field(&table_info_payload(&commit.package().exact_bytes())?)?,
            locked
        );
    }
    Ok(())
}

#[test]
fn changed_lock_preserves_unknown_wire_bytes_and_inverse_source() -> TestResult<()> {
    let source = synthetic_package(None)?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_body_table_lock(BodyTableSelector::name("Revenue"))?;
    edit.lock();
    let commit = edit.commit()?;

    assert_eq!(commit.package().body_table_lock(0usize)?, State::Locked);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        lock_field(&table_info_payload(&commit.package().exact_bytes())?)?,
        Some(true)
    );
    assert!(has_unknown_table_info_field(&table_info_payload(
        &commit.package().exact_bytes()
    )?)?);

    let restored = commit
        .package()
        .apply_body_table_lock(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    assert_eq!(restored.package().body_table_lock(0usize)?, State::Unlocked);
    assert_eq!(
        lock_field(&table_info_payload(&restored.package().exact_bytes())?)?,
        None
    );
    assert!(has_unknown_table_info_field(&table_info_payload(
        &restored.package().exact_bytes()
    )?)?);
    Ok(())
}

#[test]
fn selectors_conflicts_and_unrelated_members_are_checked() -> TestResult<()> {
    let source = synthetic_package(None)?;
    let package = Package::from_bytes(&source)?;

    assert!(matches!(
        package.body_table_lock(BodyTableSelector::name("Missing")),
        Err(BodyTableLockError::TableNotFound)
    ));
    assert!(matches!(
        package.edit_body_table_lock(BodyTableSelector::index(1)),
        Err(BodyTableLockError::TableNotFound)
    ));

    let mut edit = package.edit_body_table_lock(BodyTableSelector::name("Revenue"))?;
    edit.lock();
    let commit = edit.commit()?;
    assert_eq!(package.exact_bytes(), source.as_slice());
    assert_eq!(
        sentinel_payload(&commit.package().exact_bytes())?,
        b"untouched"
    );

    let applied = package.apply_body_table_lock(commit.patch())?;
    assert_eq!(
        applied
            .package()
            .body_table_lock(BodyTableSelector::index(0))?,
        State::Locked
    );
    assert!(matches!(
        applied.package().apply_body_table_lock(commit.patch()),
        Err(BodyTableLockError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_body_table_lock(&commit.patch().inverse()),
        Err(BodyTableLockError::PatchConflict)
    ));

    let catalog = Catalog::from_bytes(&source)?;
    let tampered_source = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"tampered-but-valid-source",
        )],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_lock(commit.patch()),
        Err(BodyTableLockError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn explicit_false_lock_round_trip_is_exact_and_semantic() -> TestResult<()> {
    let source = synthetic_package(Some(false))?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_body_table_lock(0usize)?;
    edit.lock();
    let commit = edit.commit()?;
    assert_eq!(commit.package().body_table_lock("Revenue")?, State::Locked);
    assert_eq!(
        lock_field(&table_info_payload(&commit.package().exact_bytes())?)?,
        Some(true)
    );

    let restored = commit
        .package()
        .apply_body_table_lock(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    assert_eq!(
        lock_field(&table_info_payload(&restored.package().exact_bytes())?)?,
        Some(false)
    );
    Ok(())
}

#[test]
fn legacy_table_info_message_type_round_trips_exactly() -> TestResult<()> {
    let source = rewrite_document_archive(&synthetic_package(None)?, |archive| {
        let drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        drawable.messages[message_index].type_ = LEGACY_TABLE_INFO_MESSAGE_TYPE;
        drawable.archive_info.message_infos[message_index].type_ = LEGACY_TABLE_INFO_MESSAGE_TYPE;
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_lock(BodyTableSelector::index(0))?,
        State::Unlocked
    );
    let mut edit = package.edit_body_table_lock(0usize)?;
    edit.lock();
    let commit = edit.commit()?;
    assert_eq!(commit.package().body_table_lock(0usize)?, State::Locked);
    let restored = commit
        .package()
        .apply_body_table_lock(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn noncanonical_object_length_prefix_is_rejected_before_changed_publish() -> TestResult<()> {
    let source = noncanonical_document_prefix(&synthetic_package(None)?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_lock(BodyTableSelector::index(0))?,
        State::Unlocked
    );
    let mut edit = package.edit_body_table_lock(0usize)?;
    edit.lock();
    assert!(matches!(
        edit.commit(),
        Err(BodyTableLockError::InvalidSource)
    ));
    Ok(())
}

#[test]
fn aggregate_only_native_ownership_locks_and_inverts_exactly() -> TestResult<()> {
    let source = rewrite_document_archive(&synthetic_package(None)?, |archive| {
        for (identifier, message_type) in [
            (1, ROOT_MESSAGE_TYPE),
            (BODY_IDENTIFIER, BODY_MESSAGE_TYPE),
            (ATTACHMENT_IDENTIFIER, ATTACHMENT_MESSAGE_TYPE),
            (DRAWABLE_IDENTIFIER, TABLE_INFO_MESSAGE_TYPE),
        ] {
            let object = archive.object_mut(identifier).ok_or("missing object")?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == message_type)
                .ok_or("missing message")?;
            object.archive_info.message_infos[message_index]
                .field_infos
                .clear();
        }
        let drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        drawable.archive_info.message_infos[message_index]
            .object_references
            .retain(|identifier| *identifier != BODY_IDENTIFIER);
        Ok(())
    })?;

    let package = Package::from_bytes(&source)?;
    assert_eq!(package.body_table_lock(0usize)?, State::Unlocked);
    let mut edit = package.edit_body_table_lock(0usize)?;
    edit.lock();
    let commit = edit.commit()?;
    assert_eq!(commit.package().body_table_lock(0usize)?, State::Locked);
    let restored = commit
        .package()
        .apply_body_table_lock(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn archive_header_ownership_tampering_is_rejected_by_read_and_noop_resolution() -> TestResult<()> {
    let source = synthetic_package(None)?;

    let body_missing = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(archive, BODY_IDENTIFIER, BODY_MESSAGE_TYPE, &[])
    })?;
    let attachment_missing = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(archive, ATTACHMENT_IDENTIFIER, ATTACHMENT_MESSAGE_TYPE, &[])
    })?;
    let drawable_missing = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            &[BODY_IDENTIFIER],
        )
    })?;
    let table_info_body_missing = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            &[MODEL_IDENTIFIER],
        )
    })?;
    let body_wrong = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            BODY_IDENTIFIER,
            BODY_MESSAGE_TYPE,
            &[DRAWABLE_IDENTIFIER],
        )
    })?;
    let attachment_wrong = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            ATTACHMENT_IDENTIFIER,
            ATTACHMENT_MESSAGE_TYPE,
            &[MODEL_IDENTIFIER],
        )
    })?;
    let drawable_wrong = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            &[BODY_IDENTIFIER, ATTACHMENT_IDENTIFIER],
        )
    })?;

    let body_duplicate = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            BODY_IDENTIFIER,
            BODY_MESSAGE_TYPE,
            &[ATTACHMENT_IDENTIFIER, ATTACHMENT_IDENTIFIER],
        )
    })?;
    let attachment_duplicate = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            ATTACHMENT_IDENTIFIER,
            ATTACHMENT_MESSAGE_TYPE,
            &[DRAWABLE_IDENTIFIER, DRAWABLE_IDENTIFIER],
        )
    })?;
    let drawable_duplicate = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            &[BODY_IDENTIFIER, MODEL_IDENTIFIER, MODEL_IDENTIFIER],
        )
    })?;
    let root_missing = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(archive, 1, ROOT_MESSAGE_TYPE, &[])
    })?;
    let root_wrong = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(archive, 1, ROOT_MESSAGE_TYPE, &[ATTACHMENT_IDENTIFIER])
    })?;
    let root_duplicate = rewrite_document_archive(&source, |archive| {
        rewrite_message_references(
            archive,
            1,
            ROOT_MESSAGE_TYPE,
            &[BODY_IDENTIFIER, BODY_IDENTIFIER],
        )
    })?;
    let root_field_wrong = rewrite_document_archive(&source, |archive| {
        let root = archive.object_mut(1).ok_or("missing root")?;
        let message_index = root
            .messages
            .iter()
            .position(|message| message.type_ == ROOT_MESSAGE_TYPE)
            .ok_or("missing root message")?;
        root.archive_info.message_infos[message_index].field_infos =
            vec![field_reference(vec![ROOT_BODY_FIELD + 1], BODY_IDENTIFIER)];
        Ok(())
    })?;
    let root_field_duplicate = rewrite_document_archive(&source, |archive| {
        let root = archive.object_mut(1).ok_or("missing root")?;
        let message_index = root
            .messages
            .iter()
            .position(|message| message.type_ == ROOT_MESSAGE_TYPE)
            .ok_or("missing root message")?;
        root.archive_info.message_infos[message_index]
            .field_infos
            .push(field_reference(vec![ROOT_BODY_FIELD], BODY_IDENTIFIER));
        Ok(())
    })?;
    let root_field_wrong_identifier = rewrite_document_archive(&source, |archive| {
        let root = archive.object_mut(1).ok_or("missing root")?;
        let message_index = root
            .messages
            .iter()
            .position(|message| message.type_ == ROOT_MESSAGE_TYPE)
            .ok_or("missing root message")?;
        root.archive_info.message_infos[message_index].field_infos = vec![field_reference(
            vec![ROOT_BODY_FIELD],
            ATTACHMENT_IDENTIFIER,
        )];
        Ok(())
    })?;
    let model_header_wrong = rewrite_document_archive(&source, |archive| {
        archive
            .object_mut(MODEL_IDENTIFIER)
            .ok_or("missing model")?
            .archive_info
            .identifier = Some(MODEL_IDENTIFIER + 1);
        Ok(())
    })?;
    let model_message_header_wrong = rewrite_document_archive(&source, |archive| {
        let model = archive
            .object_mut(MODEL_IDENTIFIER)
            .ok_or("missing model")?;
        let message_index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        model.messages[message_index].type_ = LEGACY_TABLE_INFO_MESSAGE_TYPE;
        Ok(())
    })?;
    let drawable_merge = rewrite_document_archive(&source, |archive| {
        archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?
            .archive_info
            .should_merge = Some(true);
        Ok(())
    })?;
    let drawable_base = rewrite_document_archive(&source, |archive| {
        rewrite_message_metadata(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            |info| info.base_message_index = Some(0),
        )
    })?;
    let drawable_diff_merge = rewrite_document_archive(&source, |archive| {
        rewrite_message_metadata(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            |info| info.diff_merge_version = vec![1],
        )
    })?;
    let drawable_diff_field = rewrite_document_archive(&source, |archive| {
        rewrite_message_metadata(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            |info| info.diff_field_path = Some(FieldPath::new(vec![1])),
        )
    })?;
    let drawable_fields_to_remove = rewrite_document_archive(&source, |archive| {
        rewrite_message_metadata(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            |info| info.fields_to_remove = vec![FieldPath::new(vec![1])],
        )
    })?;
    let drawable_diff_read = rewrite_document_archive(&source, |archive| {
        rewrite_message_metadata(
            archive,
            DRAWABLE_IDENTIFIER,
            TABLE_INFO_MESSAGE_TYPE,
            |info| info.diff_read_version = vec![1],
        )
    })?;

    for tampered in [
        body_missing,
        attachment_missing,
        drawable_missing,
        table_info_body_missing,
        body_wrong,
        attachment_wrong,
        drawable_wrong,
        body_duplicate,
        attachment_duplicate,
        drawable_duplicate,
        root_missing,
        root_wrong,
        root_duplicate,
        root_field_wrong,
        root_field_duplicate,
        root_field_wrong_identifier,
        model_header_wrong,
        model_message_header_wrong,
        drawable_merge,
        drawable_base,
        drawable_diff_merge,
        drawable_diff_field,
        drawable_fields_to_remove,
        drawable_diff_read,
    ] {
        assert_invalid_metadata_source(&tampered)?;
    }
    Ok(())
}

#[test]
fn present_model_and_attachment_field_local_ownership_is_strict() -> TestResult<()> {
    let source = synthetic_package(None)?;

    let model_empty = rewrite_document_archive(&source, |archive| {
        let drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        drawable.archive_info.message_infos[message_index].field_infos = vec![
            field_reference(
                vec![TABLE_INFO_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
                BODY_IDENTIFIER,
            ),
            FieldInfo::new(FieldPath::new(vec![TABLE_INFO_MODEL_FIELD])),
        ];
        Ok(())
    })?;

    let model_data_only = rewrite_document_archive(&source, |archive| {
        let drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        let mut model_field = FieldInfo::new(FieldPath::new(vec![TABLE_INFO_MODEL_FIELD]));
        model_field.data_references.push(MODEL_IDENTIFIER);
        drawable.archive_info.message_infos[message_index].field_infos = vec![
            field_reference(
                vec![TABLE_INFO_SUPER_FIELD, DRAWABLE_PARENT_FIELD],
                BODY_IDENTIFIER,
            ),
            model_field,
        ];
        Ok(())
    })?;

    let model_duplicate = rewrite_document_archive(&source, |archive| {
        let drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER)
            .ok_or("missing drawable")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        drawable.archive_info.message_infos[message_index]
            .field_infos
            .push(field_reference(
                vec![TABLE_INFO_MODEL_FIELD],
                MODEL_IDENTIFIER,
            ));
        Ok(())
    })?;

    let attachment_empty = rewrite_document_archive(&source, |archive| {
        let attachment = archive
            .object_mut(ATTACHMENT_IDENTIFIER)
            .ok_or("missing attachment")?;
        let message_index = attachment
            .messages
            .iter()
            .position(|message| message.type_ == ATTACHMENT_MESSAGE_TYPE)
            .ok_or("missing attachment message")?;
        attachment.archive_info.message_infos[message_index].field_infos =
            vec![FieldInfo::new(FieldPath::new(vec![DRAWABLE_FIELD]))];
        Ok(())
    })?;

    let attachment_data_only = rewrite_document_archive(&source, |archive| {
        let attachment = archive
            .object_mut(ATTACHMENT_IDENTIFIER)
            .ok_or("missing attachment")?;
        let message_index = attachment
            .messages
            .iter()
            .position(|message| message.type_ == ATTACHMENT_MESSAGE_TYPE)
            .ok_or("missing attachment message")?;
        let mut drawable_field = FieldInfo::new(FieldPath::new(vec![DRAWABLE_FIELD]));
        drawable_field.data_references.push(DRAWABLE_IDENTIFIER);
        attachment.archive_info.message_infos[message_index].field_infos = vec![drawable_field];
        Ok(())
    })?;

    let attachment_duplicate = rewrite_document_archive(&source, |archive| {
        let attachment = archive
            .object_mut(ATTACHMENT_IDENTIFIER)
            .ok_or("missing attachment")?;
        let message_index = attachment
            .messages
            .iter()
            .position(|message| message.type_ == ATTACHMENT_MESSAGE_TYPE)
            .ok_or("missing attachment message")?;
        attachment.archive_info.message_infos[message_index]
            .field_infos
            .push(field_reference(vec![DRAWABLE_FIELD], DRAWABLE_IDENTIFIER));
        Ok(())
    })?;

    for tampered in [
        model_empty,
        model_data_only,
        model_duplicate,
        attachment_empty,
        attachment_data_only,
        attachment_duplicate,
    ] {
        assert_invalid_metadata_source(&tampered)?;
    }
    Ok(())
}

#[test]
fn contradictory_and_aliased_field_local_ownership_is_rejected() -> TestResult<()> {
    let source = synthetic_package(None)?;

    let missing_aggregate = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let info = &mut body.archive_info.message_infos[message_index];
        info.object_references.clear();
        let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
        field.object_references.push(ATTACHMENT_IDENTIFIER);
        info.field_infos.push(field);
        Ok(())
    })?;

    let wrong_path = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let info = &mut body.archive_info.message_infos[message_index];
        let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD + 1]));
        field.object_references.push(ATTACHMENT_IDENTIFIER);
        info.field_infos.push(field);
        Ok(())
    })?;

    let duplicate_alias = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let info = &mut body.archive_info.message_infos[message_index];
        for _ in 0..2 {
            let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
            field.object_references.push(ATTACHMENT_IDENTIFIER);
            info.field_infos.push(field);
        }
        Ok(())
    })?;

    let contradictory_field = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let info = &mut body.archive_info.message_infos[message_index];
        let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
        field.object_references.push(MODEL_IDENTIFIER);
        info.field_infos.push(field);
        Ok(())
    })?;

    let empty_field_declaration = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        body.archive_info.message_infos[message_index].field_infos =
            vec![FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]))];
        Ok(())
    })?;

    let data_only_field_declaration = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let mut field = FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD]));
        field.data_references.push(ATTACHMENT_IDENTIFIER);
        body.archive_info.message_infos[message_index].field_infos = vec![field];
        Ok(())
    })?;

    let duplicate_empty_field_declaration = rewrite_document_archive(&source, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        body.archive_info.message_infos[message_index]
            .field_infos
            .push(FieldInfo::new(FieldPath::new(vec![TABLE_BODY_FIELD])));
        Ok(())
    })?;

    for tampered in [
        missing_aggregate,
        wrong_path,
        duplicate_alias,
        contradictory_field,
        empty_field_declaration,
        data_only_field_declaration,
        duplicate_empty_field_declaration,
    ] {
        assert_invalid_metadata_source(&tampered)?;
    }
    Ok(())
}

#[test]
fn cross_table_role_alias_is_rejected_by_read_and_noop_resolution() -> TestResult<()> {
    let source = synthetic_package_with_unique_table_count(None, 2)?;
    let tampered = rewrite_document_archive(&source, |archive| {
        let second_drawable = archive
            .object_mut(DRAWABLE_IDENTIFIER + 3)
            .ok_or("missing second drawable")?;
        let message_index = second_drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing second table-info message")?;
        let mut table_info =
            tst::TableInfoArchive::decode(second_drawable.messages[message_index].data.as_slice())?;
        table_info.table_model = reference(ATTACHMENT_IDENTIFIER);
        second_drawable.messages[message_index].data = table_info.encode_to_vec();
        second_drawable.archive_info.message_infos[message_index].object_references =
            vec![BODY_IDENTIFIER, ATTACHMENT_IDENTIFIER];
        second_drawable.archive_info.message_infos[message_index]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [TABLE_INFO_MODEL_FIELD])
            .ok_or("missing second model field")?
            .object_references = vec![ATTACHMENT_IDENTIFIER];
        Ok(())
    })?;
    assert_invalid_metadata_source(&tampered)
}

#[test]
fn body_table_lock_ignores_valid_unrelated_section_paths() -> TestResult<()> {
    const FIRST_SECTION_IDENTIFIER: u64 = 900;
    const SECOND_SECTION_IDENTIFIER: u64 = 901;

    let source = rewrite_document_archive(&synthetic_package(None)?, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let mut storage =
            tswp::StorageArchive::decode(body.messages[message_index].data.as_slice())?;
        let text = storage.text.first_mut().ok_or("missing body text")?;
        text.replace_range(..1, "\u{0004}");
        storage.table_section = Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(FIRST_SECTION_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 1,
                    object: Some(reference(SECOND_SECTION_IDENTIFIER)),
                },
            ],
        });
        body.messages[message_index].data = storage.encode_to_vec();
        let info = &mut body.archive_info.message_infos[message_index];
        info.object_references
            .extend([FIRST_SECTION_IDENTIFIER, SECOND_SECTION_IDENTIFIER]);
        let mut section_field = FieldInfo::new(FieldPath::new(vec![17]));
        section_field
            .object_references
            .extend([FIRST_SECTION_IDENTIFIER, SECOND_SECTION_IDENTIFIER]);
        info.field_infos.push(section_field);

        archive.objects.push(object(
            FIRST_SECTION_IDENTIFIER,
            10_011,
            tp::SectionArchive {
                name: Some("First".to_owned()),
                ..tp::SectionArchive::default()
            }
            .encode_to_vec(),
            &[],
        )?);
        archive.objects.push(object(
            SECOND_SECTION_IDENTIFIER,
            10_011,
            tp::SectionArchive {
                name: Some("Second".to_owned()),
                ..tp::SectionArchive::default()
            }
            .encode_to_vec(),
            &[],
        )?);
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_lock(BodyTableSelector::index(0))?,
        State::Unlocked
    );
    Ok(())
}

#[test]
fn stale_body_table_metadata_is_not_reported_as_a_missing_table() -> TestResult<()> {
    let empty = Package::from_bytes(&synthetic_package_with_table_count(None, 0)?)?;
    assert!(matches!(
        empty.body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::TableNotFound)
    ));

    let source = rewrite_document_archive(&synthetic_package(None)?, |archive| {
        let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
        let message_index = body
            .messages
            .iter()
            .position(|message| message.type_ == BODY_MESSAGE_TYPE)
            .ok_or("missing body message")?;
        let mut storage =
            tswp::StorageArchive::decode(body.messages[message_index].data.as_slice())?;
        storage.table_attachment = Some(tswp::ObjectAttributeTable { entries: vec![] });
        body.messages[message_index].data = storage.encode_to_vec();
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;

    assert!(matches!(
        package.body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    Ok(())
}

#[test]
fn missing_rooted_table_edge_fails_closed_at_the_table_lock_adapter() -> TestResult<()> {
    let source = synthetic_package(None)?;
    let catalog = Catalog::from_bytes(&source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let body = archive.object_mut(BODY_IDENTIFIER).ok_or("missing body")?;
    let body_message_index = body
        .messages
        .iter()
        .position(|message| message.type_ == BODY_MESSAGE_TYPE)
        .ok_or("missing body message")?;
    let mut storage =
        tswp::StorageArchive::decode(body.messages[body_message_index].data.as_slice())?;
    storage
        .table_attachment
        .as_mut()
        .ok_or("missing table attachment")?
        .entries
        .first_mut()
        .ok_or("missing table attachment entry")?
        .object = Some(reference(9_999));
    body.replace_message(
        body_message_index,
        RawMessage {
            type_: BODY_MESSAGE_TYPE,
            data: storage.encode_to_vec(),
        },
    )?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    let malformed = catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?;
    let package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        package.body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_body_table_lock(BodyTableSelector::index(0)),
        Err(BodyTableLockError::InvalidSource)
    ));
    Ok(())
}
