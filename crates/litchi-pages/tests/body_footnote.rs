//! Integration coverage for the selector-first Pages body-footnote lifecycle.
//!
//! The fixture deliberately carries the native metadata sidecar as well as
//! the rooted body graph.  Lifecycle edits are expected to publish one exact
//! candidate across the document and metadata components; the tests below
//! keep that contract visible while the owner implementation grows.

use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsp, tswp};
use litchi_pages::footnote::body::{Footnote, Position, Selector};
use litchi_pages::{BodyFootnoteError, Package};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const DOCUMENT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const FIRST_REFERENCE_IDENTIFIER: u64 = 100;
const FIRST_STORAGE_IDENTIFIER: u64 = 101;
const FIRST_MARKER_IDENTIFIER: u64 = 102;
const SECOND_REFERENCE_IDENTIFIER: u64 = 110;
const SECOND_STORAGE_IDENTIFIER: u64 = 111;
const SECOND_MARKER_IDENTIFIER: u64 = 112;
const VIEW_STATE_IDENTIFIER: u64 = 500;
const METADATA_IDENTIFIER: u64 = 900;
const UNRELATED_IDENTIFIER: u64 = 800;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const FOOTNOTE_CONTENT_PREFIX: &str = "\u{fffc} ";
const PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const SENTINEL_MEMBER: &str = "Data/sentinel.bin";
const LAST_OBJECT_IDENTIFIER: u64 = 1_000;
const ROOT_SAVE_TOKEN: u64 = 10;
const DOCUMENT_SAVE_TOKEN: u64 = 10;
const VIEW_STATE_SAVE_TOKEN: u64 = 7;
const UNRELATED_SAVE_TOKEN: u64 = 5;
const UNKNOWN_ROOT_FIELD: u32 = 98;
const UNKNOWN_COMPONENT_FIELD: u32 = 90;
const UNKNOWN_PAYLOAD_FIELD: u32 = 97;

struct Fixture {
    bytes: Vec<u8>,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_info(path: &[u32], references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path.to_vec());
    field.object_references.extend_from_slice(references);
    field
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn append_unknown(data: &mut Vec<u8>, field: u32, value: u64) -> TestResult {
    append_varint_field(data, field, value)?;
    Ok(())
}

fn document_payload() -> TestResult<Vec<u8>> {
    let mut payload = tp::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: litchi_iwa_protos::tsk::DocumentArchive::default(),
            view_state: Some(reference(VIEW_STATE_IDENTIFIER)),
            ..tsa::DocumentArchive::default()
        },
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    }
    .encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_PAYLOAD_FIELD, 0x0102_0304)?;
    Ok(payload)
}

fn body_payload() -> TestResult<Vec<u8>> {
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["A😀\u{e}B\u{e}C".to_owned()],
        table_footnote: Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 3,
                    object: Some(reference(FIRST_REFERENCE_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 5,
                    object: Some(reference(SECOND_REFERENCE_IDENTIFIER)),
                },
            ],
        }),
        ..tswp::StorageArchive::default()
    };
    let mut payload = body.encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_PAYLOAD_FIELD, 0x0a0b_0c0d)?;
    Ok(payload)
}

fn reference_payload(storage_identifier: u64, custom_mark: Option<&str>) -> TestResult<Vec<u8>> {
    let mut payload = tswp::FootnoteReferenceAttachmentArchive {
        super_: Some(tswp::TextualAttachmentArchive {
            string_equivalent: Some("*".to_owned()),
            kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
        }),
        contained_storage: Some(reference(storage_identifier)),
        custom_mark_string: custom_mark.map(str::to_owned),
    }
    .encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_PAYLOAD_FIELD, 0x1112_1314)?;
    Ok(payload)
}

fn storage_payload(text: &str, marker_identifier: u64) -> TestResult<Vec<u8>> {
    let mut payload = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Footnote as i32),
        text: vec![format!("{FOOTNOTE_CONTENT_PREFIX}{text}")],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: Some(reference(marker_identifier)),
            }],
        }),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_PAYLOAD_FIELD, 0x1516_1718)?;
    Ok(payload)
}

fn marker_payload() -> TestResult<Vec<u8>> {
    let mut payload = tswp::TextualAttachmentArchive {
        string_equivalent: Some("*".to_owned()),
        kind: Some(tswp::textual_attachment_archive::Kind::KKindFootnoteMark as i32),
    }
    .encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_PAYLOAD_FIELD, 0x191a_1b1c)?;
    Ok(payload)
}

fn add_graph_metadata(object: &mut ArchiveObject, references: &[u64], fields: &[(&[u32], &[u64])]) {
    let info = &mut object.archive_info.message_infos[0];
    info.object_references.extend_from_slice(references);
    for &(path, field_references) in fields {
        info.field_infos.push(field_info(path, field_references));
    }
}

fn uuid_entry(identifier: u64, lower: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower,
            upper: lower.saturating_add(10_000),
        },
    }
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    uuid_entries: &[tsp::ObjectUuidMapEntry],
) -> TestResult<Vec<u8>> {
    let component = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        object_uuid_map_entries: uuid_entries.to_vec(),
        save_token: Some(token),
        ..tsp::ComponentInfo::default()
    };
    let mut payload = component.encode_to_vec();
    append_unknown(&mut payload, UNKNOWN_COMPONENT_FIELD, identifier + 0x9000)?;
    Ok(payload)
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let mut payload = tsp::PackageMetadata {
        last_object_identifier: LAST_OBJECT_IDENTIFIER,
        save_token: Some(ROOT_SAVE_TOKEN),
        components: Vec::new(),
        ..tsp::PackageMetadata::default()
    }
    .encode_to_vec();
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(
            DOCUMENT_IDENTIFIER,
            "Document",
            DOCUMENT_SAVE_TOKEN,
            &[
                uuid_entry(FIRST_STORAGE_IDENTIFIER, 1),
                uuid_entry(SECOND_STORAGE_IDENTIFIER, 2),
            ],
        )?,
    )?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(2, "ViewState", VIEW_STATE_SAVE_TOKEN, &[])?,
    )?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(3, "Unrelated", UNRELATED_SAVE_TOKEN, &[])?,
    )?;
    append_unknown(&mut payload, UNKNOWN_ROOT_FIELD, 0xfeed_beef)?;
    Ok(payload)
}

fn build_fixture() -> TestResult<Fixture> {
    let mut root = object(
        DOCUMENT_IDENTIFIER,
        DOCUMENT_MESSAGE_TYPE,
        document_payload()?,
    )?;
    add_graph_metadata(
        &mut root,
        &[BODY_IDENTIFIER, VIEW_STATE_IDENTIFIER],
        &[
            (&[4], &[BODY_IDENTIFIER]),
            (&[15, 5], &[VIEW_STATE_IDENTIFIER]),
        ],
    );

    let mut body = object(BODY_IDENTIFIER, STORAGE_MESSAGE_TYPE, body_payload()?)?;
    add_graph_metadata(
        &mut body,
        &[FIRST_REFERENCE_IDENTIFIER, SECOND_REFERENCE_IDENTIFIER],
        &[(
            &[16, 1, 2],
            &[FIRST_REFERENCE_IDENTIFIER, SECOND_REFERENCE_IDENTIFIER],
        )],
    );

    let mut first_reference = object(
        FIRST_REFERENCE_IDENTIFIER,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        reference_payload(FIRST_STORAGE_IDENTIFIER, None)?,
    )?;
    add_graph_metadata(
        &mut first_reference,
        &[FIRST_STORAGE_IDENTIFIER],
        &[(&[2], &[FIRST_STORAGE_IDENTIFIER])],
    );
    let mut first_storage = object(
        FIRST_STORAGE_IDENTIFIER,
        STORAGE_MESSAGE_TYPE,
        storage_payload("First", FIRST_MARKER_IDENTIFIER)?,
    )?;
    add_graph_metadata(
        &mut first_storage,
        &[FIRST_MARKER_IDENTIFIER],
        &[(&[9, 1, 2], &[FIRST_MARKER_IDENTIFIER])],
    );
    let first_marker = object(
        FIRST_MARKER_IDENTIFIER,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        marker_payload()?,
    )?;

    let mut second_reference = object(
        SECOND_REFERENCE_IDENTIFIER,
        FOOTNOTE_REFERENCE_MESSAGE_TYPE,
        reference_payload(SECOND_STORAGE_IDENTIFIER, Some("†"))?,
    )?;
    add_graph_metadata(
        &mut second_reference,
        &[SECOND_STORAGE_IDENTIFIER],
        &[(&[2], &[SECOND_STORAGE_IDENTIFIER])],
    );
    let mut second_storage = object(
        SECOND_STORAGE_IDENTIFIER,
        STORAGE_MESSAGE_TYPE,
        storage_payload("Second", SECOND_MARKER_IDENTIFIER)?,
    )?;
    add_graph_metadata(
        &mut second_storage,
        &[SECOND_MARKER_IDENTIFIER],
        &[(&[9, 1, 2], &[SECOND_MARKER_IDENTIFIER])],
    );
    let second_marker = object(
        SECOND_MARKER_IDENTIFIER,
        TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
        marker_payload()?,
    )?;

    let document_component = compressed(vec![
        root,
        body,
        first_reference,
        first_storage,
        first_marker,
        second_reference,
        second_storage,
        second_marker,
    ])?;
    let metadata_component = compressed(vec![object(
        METADATA_IDENTIFIER,
        METADATA_MESSAGE_TYPE,
        metadata_payload()?,
    )?])?;
    let view_state_component = compressed(vec![object(
        VIEW_STATE_IDENTIFIER,
        VIEW_STATE_MESSAGE_TYPE,
        b"view-state-preservation-witness".to_vec(),
    )?])?;
    let unrelated_component = compressed(vec![object(
        UNRELATED_IDENTIFIER,
        999,
        b"unrelated-component-preservation-witness".to_vec(),
    )?])?;

    let entries = [
        (DOCUMENT_MEMBER, document_component.as_slice()),
        (METADATA_MEMBER, metadata_component.as_slice()),
        (VIEW_STATE_MEMBER, view_state_component.as_slice()),
        (UNRELATED_MEMBER, unrelated_component.as_slice()),
        (
            SENTINEL_MEMBER,
            b"zip-sentinel-preservation-witness".as_slice(),
        ),
        (PREVIEW_NAMES[0], b"preview-full".as_slice()),
        (PREVIEW_NAMES[1], b"preview-micro".as_slice()),
        (PREVIEW_NAMES[2], b"preview-web".as_slice()),
    ];
    Ok(Fixture {
        bytes: litchi_iwa_archive::package::to_bytes(entries, Limits::default())?,
    })
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn metadata_payload_from_bytes(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("metadata component is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    archive
        .objects
        .first()
        .and_then(|object| object.messages.first())
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("metadata object is missing").into())
}

fn metadata(bytes: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        metadata_payload_from_bytes(bytes)?.as_slice(),
    )?)
}

fn component(metadata: &tsp::PackageMetadata, identifier: u64) -> TestResult<&tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| component.identifier == identifier)
        .ok_or_else(|| io::Error::other("metadata component is missing").into())
}

fn object_ids(bytes: &[u8], member: &str) -> TestResult<Vec<u64>> {
    let catalog = Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("component is missing"))?;
    Ok(
        Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?
            .objects
            .into_iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect(),
    )
}

fn has_entry(bytes: &[u8], name: &str) -> TestResult<bool> {
    Ok(Catalog::from_bytes(bytes)?
        .iter()
        .any(|entry| entry.name() == name))
}

fn rewrite_member(
    source: &[u8],
    member: &str,
    edit: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("component is missing"))?;
    let mut archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    edit(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(member, compressed.as_slice())],
        Limits::default(),
    )?)
}

fn rewrite_object_payload(
    source: &[u8],
    member: &str,
    identifier: u64,
    edit: impl FnOnce(u32, Vec<u8>) -> TestResult<(u32, Vec<u8>)>,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, member, |archive| {
        let object = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(identifier))
            .ok_or_else(|| io::Error::other("object is missing"))?;
        let message = object
            .messages
            .first()
            .ok_or_else(|| io::Error::other("message is missing"))?;
        let (type_, data) = edit(message.type_, message.data.clone())?;
        object.replace_message_preserving_header(0, RawMessage { type_, data })?;
        Ok(())
    })
}

fn without_member(source: &[u8], member: &str) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
            &[],
            &[member],
            Limits::default(),
        )?,
    )
}

fn append_versioned_external_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_from_bytes(source)?;
    let component = tsp::ComponentInfo {
        identifier: 2,
        preferred_locator: "ViewState".to_owned(),
        locator: Some("ViewState".to_owned()),
        external_references: vec![tsp::ComponentExternalReference {
            component_identifier: DOCUMENT_IDENTIFIER,
            object_identifier: Some(LAST_OBJECT_IDENTIFIER + 2),
            is_weak: Some(true),
        }],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, 11, &component)?;
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn add_metadata_uuid_alias(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut decoded =
        tsp::PackageMetadata::decode(metadata_payload_from_bytes(source)?.as_slice())?;
    let document = decoded
        .components
        .iter_mut()
        .find(|component| component.identifier == DOCUMENT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("document metadata component is missing"))?;
    document
        .object_uuid_map_entries
        .push(uuid_entry(identifier, 99));
    let payload = decoded.encode_to_vec();
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn remove_metadata_uuid_registration(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut decoded =
        tsp::PackageMetadata::decode(metadata_payload_from_bytes(source)?.as_slice())?;
    let document = decoded
        .components
        .iter_mut()
        .find(|component| component.identifier == DOCUMENT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("document metadata component is missing"))?;
    document
        .object_uuid_map_entries
        .retain(|entry| entry.identifier != identifier);
    let payload = decoded.encode_to_vec();
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn duplicate_body_table(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(source, DOCUMENT_MEMBER, BODY_IDENTIFIER, |_type_, data| {
        let body = tswp::StorageArchive::decode(data.as_slice())?;
        let table = body
            .table_footnote
            .clone()
            .ok_or_else(|| io::Error::other("body footnote table is missing"))?;
        let mut payload = body.encode_to_vec();
        append_length_delimited_field(&mut payload, 16, &table.encode_to_vec())?;
        Ok((STORAGE_MESSAGE_TYPE, payload))
    })
}

fn add_ambiguous_document_component(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_from_bytes(source)?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(DOCUMENT_IDENTIFIER, "Document", DOCUMENT_SAVE_TOKEN, &[])?,
    )?;
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn remove_document_metadata_component(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut decoded =
        tsp::PackageMetadata::decode(metadata_payload_from_bytes(source)?.as_slice())?;
    decoded
        .components
        .retain(|component| component.identifier != DOCUMENT_IDENTIFIER);
    let payload = decoded.encode_to_vec();
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn missing_body_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(source, DOCUMENT_MEMBER, BODY_IDENTIFIER, |_type_, data| {
        let mut body = tswp::StorageArchive::decode(data.as_slice())?;
        body.table_footnote
            .as_mut()
            .ok_or_else(|| io::Error::other("body footnote table is missing"))?
            .entries[0]
            .object = Some(reference(9_999));
        Ok((STORAGE_MESSAGE_TYPE, body.encode_to_vec()))
    })
}

fn deprecated_body_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(source, DOCUMENT_MEMBER, BODY_IDENTIFIER, |_type_, data| {
        let mut body = tswp::StorageArchive::decode(data.as_slice())?;
        body.table_footnote
            .as_mut()
            .ok_or_else(|| io::Error::other("body footnote table is missing"))?
            .entries[0]
            .object
            .as_mut()
            .ok_or_else(|| io::Error::other("body reference is missing"))?
            .deprecated_type = Some(17);
        Ok((STORAGE_MESSAGE_TYPE, body.encode_to_vec()))
    })
}

fn shared_storage_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_REFERENCE_IDENTIFIER,
        |_type_, data| {
            let mut attachment = tswp::FootnoteReferenceAttachmentArchive::decode(data.as_slice())?;
            attachment.contained_storage = Some(reference(SECOND_STORAGE_IDENTIFIER));
            Ok((FOOTNOTE_REFERENCE_MESSAGE_TYPE, attachment.encode_to_vec()))
        },
    )
}

fn shared_marker_attachment(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_STORAGE_IDENTIFIER,
        |_type_, data| {
            let mut storage = tswp::StorageArchive::decode(data.as_slice())?;
            storage
                .table_attachment
                .as_mut()
                .ok_or_else(|| io::Error::other("attachment table is missing"))?
                .entries[0]
                .object = Some(reference(SECOND_MARKER_IDENTIFIER));
            Ok((STORAGE_MESSAGE_TYPE, storage.encode_to_vec()))
        },
    )
}

fn extra_marker_attachment(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_STORAGE_IDENTIFIER,
        |_type_, data| {
            let mut storage = tswp::StorageArchive::decode(data.as_slice())?;
            storage
                .table_attachment
                .as_mut()
                .ok_or_else(|| io::Error::other("attachment table is missing"))?
                .entries
                .push(tswp::object_attribute_table::ObjectAttribute {
                    character_index: 1,
                    object: Some(reference(SECOND_MARKER_IDENTIFIER)),
                });
            Ok((STORAGE_MESSAGE_TYPE, storage.encode_to_vec()))
        },
    )
}

fn typed_reference_alias(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_REFERENCE_IDENTIFIER,
        |_type_, data| Ok((STORAGE_MESSAGE_TYPE, data)),
    )
}

fn duplicate_body_aggregate_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let body = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(BODY_IDENTIFIER))
            .ok_or_else(|| io::Error::other("body object is missing"))?;
        body.archive_info.message_infos[0]
            .object_references
            .push(FIRST_REFERENCE_IDENTIFIER);
        Ok(())
    })
}

fn duplicate_body_field_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let body = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(BODY_IDENTIFIER))
            .ok_or_else(|| io::Error::other("body object is missing"))?;
        body.archive_info.message_infos[0]
            .field_infos
            .push(field_info(&[16, 1, 2], &[FIRST_REFERENCE_IDENTIFIER]));
        Ok(())
    })
}

fn duplicate_body_data_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let body = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(BODY_IDENTIFIER))
            .ok_or_else(|| io::Error::other("body object is missing"))?;
        body.archive_info.message_infos[0]
            .data_references
            .push(FIRST_STORAGE_IDENTIFIER);
        Ok(())
    })
}

fn duplicate_current_external_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut decoded =
        tsp::PackageMetadata::decode(metadata_payload_from_bytes(source)?.as_slice())?;
    let view_state = decoded
        .components
        .iter_mut()
        .find(|component| component.identifier == 2)
        .ok_or_else(|| io::Error::other("view-state metadata component is missing"))?;
    view_state
        .external_references
        .push(tsp::ComponentExternalReference {
            component_identifier: DOCUMENT_IDENTIFIER,
            object_identifier: Some(LAST_OBJECT_IDENTIFIER + 2),
            is_weak: Some(true),
        });
    let payload = decoded.encode_to_vec();
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn missing_marker_object(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(FIRST_MARKER_IDENTIFIER));
        Ok(())
    })
}

fn versioned_metadata_external_reference(source: &[u8]) -> TestResult<Vec<u8>> {
    append_versioned_external_reference(source)
}

fn without_view_state(source: &[u8]) -> TestResult<Vec<u8>> {
    let rootless = rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        DOCUMENT_IDENTIFIER,
        |_type_, data| {
            let mut root = tp::DocumentArchive::decode(data.as_slice())?;
            root.super_.view_state = None;
            Ok((DOCUMENT_MESSAGE_TYPE, root.encode_to_vec()))
        },
    )?;
    let mut decoded =
        tsp::PackageMetadata::decode(metadata_payload_from_bytes(&rootless)?.as_slice())?;
    decoded
        .components
        .retain(|component| component.identifier != 2);
    let payload = decoded.encode_to_vec();
    let without_component = rewrite_object_payload(
        &rootless,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )?;
    without_member(&without_component, VIEW_STATE_MEMBER)
}

fn assert_unchanged(package: &Package, source: &[u8]) -> TestResult {
    assert_eq!(exact_bytes(package)?, source);
    Ok(())
}

fn assert_metadata_unknowns(bytes: &[u8]) -> TestResult {
    let payload = metadata_payload_from_bytes(bytes)?;
    let view = WireView::parse(&payload)?;
    assert!(
        view.fields()
            .any(|field| field.number() == UNKNOWN_ROOT_FIELD)
    );
    for field in view.fields().filter(|field| field.number() == 3) {
        let nested = field.payload();
        assert!(
            WireView::parse(nested)?
                .fields()
                .any(|nested_field| nested_field.number() == UNKNOWN_COMPONENT_FIELD)
        );
    }
    Ok(())
}

fn assert_graph_unknowns(bytes: &[u8]) -> TestResult {
    let catalog = Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("document component is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    for identifier in [
        BODY_IDENTIFIER,
        FIRST_REFERENCE_IDENTIFIER,
        FIRST_STORAGE_IDENTIFIER,
        FIRST_MARKER_IDENTIFIER,
        SECOND_REFERENCE_IDENTIFIER,
        SECOND_STORAGE_IDENTIFIER,
        SECOND_MARKER_IDENTIFIER,
    ] {
        let object = archive
            .objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(identifier))
            .ok_or_else(|| io::Error::other("graph object is missing"))?;
        let message = object
            .messages
            .first()
            .ok_or_else(|| io::Error::other("graph message is missing"))?;
        assert!(
            WireView::parse(&message.data)?
                .fields()
                .any(|field| field.number() == UNKNOWN_PAYLOAD_FIELD)
        );
    }
    Ok(())
}

#[test]
fn rich_fixture_reads_two_notes_and_preserves_metadata_identity_contract() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let notes = package.body_footnotes()?;
    assert_eq!(notes.len(), 2);
    assert_eq!(notes[0].position, Position::from_utf16_index(3)?);
    assert_eq!(notes[0].text.as_ref(), "First");
    assert_eq!(notes[0].custom_mark, None);
    assert_eq!(notes[1].position, Position::from_utf16_index(5)?);
    assert_eq!(notes[1].text.as_ref(), "Second");
    assert_eq!(notes[1].custom_mark.as_deref(), Some("†"));
    let metadata = metadata(&fixture.bytes)?;
    let document = component(&metadata, DOCUMENT_IDENTIFIER)?;
    assert_eq!(metadata.last_object_identifier, LAST_OBJECT_IDENTIFIER);
    assert_eq!(metadata.save_token, Some(ROOT_SAVE_TOKEN));
    assert_eq!(document.save_token, Some(DOCUMENT_SAVE_TOKEN));
    assert_eq!(document.object_uuid_map_entries.len(), 2);
    assert_eq!(
        document
            .object_uuid_map_entries
            .iter()
            .map(|entry| entry.identifier)
            .collect::<Vec<_>>(),
        vec![FIRST_STORAGE_IDENTIFIER, SECOND_STORAGE_IDENTIFIER]
    );
    assert_metadata_unknowns(&fixture.bytes)?;
    assert_graph_unknowns(&fixture.bytes)?;
    Ok(())
}

#[test]
fn semantic_snapshot_is_selector_first_and_reuses_one_bounded_read() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let snapshot = package.body_footnote_snapshot()?;

    assert_eq!(snapshot.len(), 2);
    assert!(!snapshot.is_empty());
    assert_eq!(
        snapshot.get(0).map(|note| note.text.as_ref()),
        Some("First")
    );
    assert_eq!(
        snapshot.select(1usize).map(|note| note.text.as_ref()),
        Some("Second")
    );
    assert_eq!(
        snapshot
            .select(Position::from_utf16_index(5)?)
            .map(|note| note.custom_mark.as_deref()),
        Some(Some("†"))
    );
    assert!(snapshot.select(Selector::index(2)).is_none());
    assert_eq!(snapshot.iter().count(), snapshot.len());
    assert_eq!(snapshot.as_slice(), package.body_footnotes()?.as_slice());

    let cloned = snapshot.clone();
    assert_eq!(cloned.as_slice(), snapshot.as_slice());
    let edit = package.edit_body_footnote(1usize)?;
    assert_eq!(edit.before(), snapshot.get(1).expect("second footnote"));
    Ok(())
}

#[test]
fn insert_publishes_graph_metadata_tokens_external_reference_and_inverse() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let position = Position::from_utf16_index(1)?;
    let commit = package.insert_body_footnote(position, "Inserted", Some("‡"))?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    let target = exact_bytes(commit.package())?;
    assert_ne!(target, fixture.bytes);
    assert!(!has_entry(&target, PREVIEW_NAMES[0])?);
    assert!(!has_entry(&target, PREVIEW_NAMES[1])?);
    assert!(!has_entry(&target, PREVIEW_NAMES[2])?);
    assert!(has_entry(&target, SENTINEL_MEMBER)?);
    assert!(has_entry(&target, UNRELATED_MEMBER)?);
    let notes = commit.package().body_footnotes()?;
    assert_eq!(
        notes[0],
        Footnote::with_custom_mark(position, "Inserted", Some("‡".into()))?
    );
    assert_eq!(notes[1].position, Position::from_utf16_index(4)?);
    assert_eq!(notes[2].position, Position::from_utf16_index(6)?);

    let target_metadata = metadata(&target)?;
    assert_eq!(
        target_metadata.last_object_identifier,
        LAST_OBJECT_IDENTIFIER + 3
    );
    assert_eq!(target_metadata.save_token, Some(ROOT_SAVE_TOKEN + 1));
    assert_eq!(
        component(&target_metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(DOCUMENT_SAVE_TOKEN + 1)
    );
    assert_eq!(
        component(&target_metadata, 2)?.save_token,
        Some(ROOT_SAVE_TOKEN + 1)
    );
    assert_eq!(
        component(&target_metadata, 3)?.save_token,
        Some(UNRELATED_SAVE_TOKEN)
    );
    let document = component(&target_metadata, DOCUMENT_IDENTIFIER)?;
    let added = document
        .object_uuid_map_entries
        .iter()
        .find(|entry| entry.identifier > SECOND_STORAGE_IDENTIFIER)
        .ok_or_else(|| io::Error::other("inserted storage UUID is missing"))?;
    let new_storage = added.identifier;
    assert_eq!(new_storage, LAST_OBJECT_IDENTIFIER + 2);
    let target_ids = object_ids(&target, DOCUMENT_MEMBER)?;
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 1)));
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 2)));
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 3)));
    assert_eq!(
        target_metadata.last_object_identifier,
        LAST_OBJECT_IDENTIFIER + 3
    );
    assert!(!document
        .object_uuid_map_entries
        .iter()
        .any(|entry| entry.identifier == new_storage - 1 || entry.identifier == new_storage + 1));
    let view_state = component(&target_metadata, 2)?;
    assert!(view_state.external_references.iter().any(|reference| {
        reference.component_identifier == DOCUMENT_IDENTIFIER
            && reference.object_identifier == Some(new_storage)
            && reference.is_weak == Some(true)
    }));
    assert_metadata_unknowns(&target)?;
    assert_graph_unknowns(&target)?;

    let forward = package.apply_body_footnote(commit.patch())?;
    assert_eq!(exact_bytes(forward.package())?, target);
    let inverse = commit
        .package()
        .apply_body_footnote(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(inverse.package())?, fixture.bytes);
    assert_eq!(commit.patch().inverse().inverse(), commit.patch().clone());
    Ok(())
}

#[test]
fn edit_set_custom_mark_and_clear_inserted_identity_atomically() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let inserted =
        package.insert_body_footnote(Position::from_utf16_index(1)?, "Inserted", None)?;
    let mut edit = inserted.package().edit_body_footnote(Selector::index(0))?;
    edit.set("Updated")?.set_custom_mark(Some("§"))?;
    let changed = edit.commit()?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    let notes = changed.package().body_footnotes()?;
    assert_eq!(notes[0].text.as_ref(), "Updated");
    assert_eq!(notes[0].custom_mark.as_deref(), Some("§"));
    assert_eq!(notes[1].text.as_ref(), "First");
    assert_eq!(notes[2].text.as_ref(), "Second");
    let changed_bytes = exact_bytes(changed.package())?;
    assert_metadata_unknowns(&changed_bytes)?;
    assert_graph_unknowns(&changed_bytes)?;

    let mut remove = changed.package().edit_body_footnote(Selector::index(0))?;
    let removed_before = remove.before().clone();
    remove.clear();
    let removed = remove.commit()?;
    assert_eq!(removed.diagnostics().touched_components(), 1);
    assert_eq!(removed.patch().before(), Some(&removed_before));
    assert_eq!(removed.patch().after(), None);
    let remaining = removed.package().body_footnotes()?;
    assert_eq!(remaining.len(), 2);
    assert_eq!(remaining[0].position, Position::from_utf16_index(3)?);
    assert_eq!(remaining[0].text.as_ref(), "First");
    assert_eq!(remaining[1].position, Position::from_utf16_index(5)?);
    assert_eq!(remaining[1].text.as_ref(), "Second");
    let target = exact_bytes(removed.package())?;
    let target_metadata = metadata(&target)?;
    assert_eq!(
        target_metadata.last_object_identifier,
        LAST_OBJECT_IDENTIFIER + 3
    );
    assert_eq!(target_metadata.save_token, Some(ROOT_SAVE_TOKEN + 2));
    let document = component(&target_metadata, DOCUMENT_IDENTIFIER)?;
    assert_eq!(document.save_token, Some(DOCUMENT_SAVE_TOKEN + 2));
    assert_eq!(document.object_uuid_map_entries.len(), 2);
    assert!(
        document
            .object_uuid_map_entries
            .iter()
            .all(|entry| entry.identifier == FIRST_STORAGE_IDENTIFIER
                || entry.identifier == SECOND_STORAGE_IDENTIFIER)
    );
    assert_eq!(
        component(&target_metadata, 2)?.save_token,
        Some(ROOT_SAVE_TOKEN + 2)
    );
    assert!(
        component(&target_metadata, 2)?
            .external_references
            .is_empty()
    );
    assert!(object_ids(&target, DOCUMENT_MEMBER)?.contains(&FIRST_REFERENCE_IDENTIFIER));
    assert!(object_ids(&target, DOCUMENT_MEMBER)?.contains(&FIRST_STORAGE_IDENTIFIER));
    assert!(object_ids(&target, DOCUMENT_MEMBER)?.contains(&FIRST_MARKER_IDENTIFIER));
    assert!(!object_ids(&target, DOCUMENT_MEMBER)?.contains(&(LAST_OBJECT_IDENTIFIER + 1)));
    assert!(!object_ids(&target, DOCUMENT_MEMBER)?.contains(&(LAST_OBJECT_IDENTIFIER + 2)));
    assert!(!object_ids(&target, DOCUMENT_MEMBER)?.contains(&(LAST_OBJECT_IDENTIFIER + 3)));
    assert_metadata_unknowns(&target)?;
    assert_graph_unknowns(&target)?;

    let restored = removed
        .package()
        .apply_body_footnote(&removed.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(changed.package())?
    );
    Ok(())
}

#[test]
fn remove_middle_note_preserves_inserted_and_following_identity() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let inserted =
        package.insert_body_footnote(Position::from_utf16_index(1)?, "Inserted", None)?;
    let mut remove = inserted.package().edit_body_footnote(Selector::index(1))?;
    let removed_before = remove.before().clone();
    assert_eq!(removed_before.text.as_ref(), "First");
    remove.clear();
    let removed = remove.commit()?;
    assert_eq!(removed.diagnostics().touched_components(), 1);
    assert_eq!(removed.patch().before(), Some(&removed_before));
    assert_eq!(removed.patch().after(), None);

    let remaining = removed.package().body_footnotes()?;
    assert_eq!(remaining.len(), 2);
    assert_eq!(remaining[0].position, Position::from_utf16_index(1)?);
    assert_eq!(remaining[0].text.as_ref(), "Inserted");
    assert_eq!(remaining[0].custom_mark, None);
    assert_eq!(remaining[1].position, Position::from_utf16_index(5)?);
    assert_eq!(remaining[1].text.as_ref(), "Second");
    assert_eq!(remaining[1].custom_mark.as_deref(), Some("†"));

    let target = exact_bytes(removed.package())?;
    let target_metadata = metadata(&target)?;
    assert_eq!(
        target_metadata.last_object_identifier,
        LAST_OBJECT_IDENTIFIER + 3
    );
    assert_eq!(target_metadata.save_token, Some(ROOT_SAVE_TOKEN + 2));
    assert_eq!(
        component(&target_metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(ROOT_SAVE_TOKEN + 2)
    );
    assert_eq!(
        component(&target_metadata, 2)?.save_token,
        Some(ROOT_SAVE_TOKEN + 1)
    );
    let document = component(&target_metadata, DOCUMENT_IDENTIFIER)?;
    let ids = document
        .object_uuid_map_entries
        .iter()
        .map(|entry| entry.identifier)
        .collect::<Vec<_>>();
    assert_eq!(ids.len(), 2);
    assert!(ids.contains(&SECOND_STORAGE_IDENTIFIER));
    assert!(ids.contains(&(LAST_OBJECT_IDENTIFIER + 2)));
    assert!(!ids.contains(&FIRST_STORAGE_IDENTIFIER));
    let view_state = component(&target_metadata, 2)?;
    assert!(view_state.external_references.iter().any(|reference| {
        reference.component_identifier == DOCUMENT_IDENTIFIER
            && reference.object_identifier == Some(LAST_OBJECT_IDENTIFIER + 2)
            && reference.is_weak == Some(true)
    }));

    let target_ids = object_ids(&target, DOCUMENT_MEMBER)?;
    assert!(!target_ids.contains(&FIRST_REFERENCE_IDENTIFIER));
    assert!(!target_ids.contains(&FIRST_STORAGE_IDENTIFIER));
    assert!(!target_ids.contains(&FIRST_MARKER_IDENTIFIER));
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 1)));
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 2)));
    assert!(target_ids.contains(&(LAST_OBJECT_IDENTIFIER + 3)));
    assert!(target_ids.contains(&SECOND_REFERENCE_IDENTIFIER));
    assert!(target_ids.contains(&SECOND_STORAGE_IDENTIFIER));
    assert!(target_ids.contains(&SECOND_MARKER_IDENTIFIER));
    assert_metadata_unknowns(&target)?;

    let restored = removed
        .package()
        .apply_body_footnote(&removed.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(inserted.package())?
    );
    Ok(())
}

#[test]
fn no_op_and_patch_conflicts_are_exact_and_side_effect_free() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let mut edit = package.edit_body_footnote(Selector::index(1))?;
    edit.set("Second")?.set_custom_mark(Some("†"))?;
    let noop = edit.commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_unchanged(noop.package(), &fixture.bytes)?;

    let commit = package.insert_body_footnote(Position::from_utf16_index(1)?, "x", None)?;
    let source_before = exact_bytes(&package)?;
    let error = package.apply_body_footnote(&commit.patch().inverse());
    assert!(matches!(error, Err(BodyFootnoteError::PatchConflict)));
    assert_eq!(exact_bytes(&package)?, source_before);
    let candidate = commit.package();
    let candidate_before = exact_bytes(candidate)?;
    let error = candidate.apply_body_footnote(commit.patch());
    assert!(matches!(error, Err(BodyFootnoteError::PatchConflict)));
    assert_eq!(exact_bytes(candidate)?, candidate_before);
    Ok(())
}

#[test]
fn insertion_rejects_boundaries_markers_occupancy_and_large_values_atomically() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let source = exact_bytes(&package)?;

    let split_astral = package.insert_body_footnote(Position::from_utf16_index(2)?, "x", None);
    assert!(matches!(
        split_astral,
        Err(BodyFootnoteError::PositionOutOfBounds)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let occupied = package.insert_body_footnote(Position::from_utf16_index(3)?, "x", None);
    assert!(matches!(occupied, Err(BodyFootnoteError::PositionOccupied)));
    assert_eq!(exact_bytes(&package)?, source);

    for value in ["bad\u{e}", "bad\u{fffc}"] {
        let text_error = package.insert_body_footnote(Position::ZERO, value, None);
        assert!(matches!(
            text_error,
            Err(BodyFootnoteError::StructuralMarker)
        ));
        let mark_error = package.insert_body_footnote(Position::ZERO, "ok", Some(value));
        assert!(matches!(
            mark_error,
            Err(BodyFootnoteError::StructuralMarker)
        ));
        assert_eq!(exact_bytes(&package)?, source);
    }
    let too_large = "x".repeat(16 * 1024 * 1024 + 1);
    assert!(matches!(
        package.insert_body_footnote(Position::ZERO, "ok", Some(&too_large)),
        Err(BodyFootnoteError::CustomMarkTooLarge)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn edit_rejects_structural_marker_and_preserves_source() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let source = exact_bytes(&package)?;
    let mut edit = package.edit_body_footnote(Selector::index(0))?;
    assert!(matches!(
        edit.set("bad\u{e}"),
        Err(BodyFootnoteError::StructuralMarker)
    ));
    assert!(matches!(
        edit.set_custom_mark(Some("bad\u{fffc}")),
        Err(BodyFootnoteError::StructuralMarker)
    ));
    assert_unchanged(&package, &source)?;
    Ok(())
}

fn assert_insert_graph_rejection(bytes: Vec<u8>) -> TestResult {
    let package = match Package::from_bytes(&bytes) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let source = exact_bytes(&package)?;
    let result = package.insert_body_footnote(Position::ZERO, "new", None);
    assert!(matches!(
        result,
        Err(BodyFootnoteError::InvalidSource
            | BodyFootnoteError::UnsupportedDependency
            | BodyFootnoteError::UnsupportedSource)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn malformed_and_aliased_graphs_reject_before_publication() -> TestResult {
    let fixture = build_fixture()?;
    for malformed in [
        duplicate_body_table(&fixture.bytes)?,
        missing_body_reference(&fixture.bytes)?,
        deprecated_body_reference(&fixture.bytes)?,
        shared_storage_reference(&fixture.bytes)?,
        shared_marker_attachment(&fixture.bytes)?,
        extra_marker_attachment(&fixture.bytes)?,
        typed_reference_alias(&fixture.bytes)?,
        duplicate_body_aggregate_reference(&fixture.bytes)?,
        duplicate_body_field_reference(&fixture.bytes)?,
        duplicate_body_data_reference(&fixture.bytes)?,
        missing_marker_object(&fixture.bytes)?,
        add_metadata_uuid_alias(&fixture.bytes, FIRST_REFERENCE_IDENTIFIER)?,
        add_metadata_uuid_alias(&fixture.bytes, FIRST_MARKER_IDENTIFIER)?,
        add_ambiguous_document_component(&fixture.bytes)?,
        remove_document_metadata_component(&fixture.bytes)?,
        without_member(&fixture.bytes, METADATA_MEMBER)?,
    ] {
        assert_insert_graph_rejection(malformed)?;
    }
    Ok(())
}

#[test]
fn missing_storage_uuid_registration_rejects_insert_and_remove_atomically() -> TestResult {
    let fixture = build_fixture()?;
    let hostile = remove_metadata_uuid_registration(&fixture.bytes, FIRST_STORAGE_IDENTIFIER)?;
    assert!(has_entry(&hostile, METADATA_MEMBER)?);
    let package = Package::from_bytes(&hostile)?;
    let source = exact_bytes(&package)?;

    let insertion = package.insert_body_footnote(Position::ZERO, "new", None);
    assert!(matches!(
        insertion,
        Err(BodyFootnoteError::InvalidSource
            | BodyFootnoteError::UnsupportedDependency
            | BodyFootnoteError::UnsupportedSource)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let mut removal = package.edit_body_footnote(Selector::index(0))?;
    removal.clear();
    let removal = removal.commit();
    assert!(matches!(
        removal,
        Err(BodyFootnoteError::InvalidSource
            | BodyFootnoteError::UnsupportedDependency
            | BodyFootnoteError::UnsupportedSource)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn versioned_metadata_external_owner_rejects_remove_atomically() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let inserted = package.insert_body_footnote(Position::ZERO, "new", None)?;
    let inserted_bytes = exact_bytes(inserted.package())?;
    let hostile = versioned_metadata_external_reference(&inserted_bytes)?;
    let hostile_package = Package::from_bytes(&hostile)?;
    let source = exact_bytes(&hostile_package)?;
    let mut edit = hostile_package.edit_body_footnote(Selector::index(0))?;
    edit.clear();
    let result = edit.commit();
    assert!(
        matches!(
            &result,
            Err(BodyFootnoteError::InvalidSource
                | BodyFootnoteError::UnsupportedDependency
                | BodyFootnoteError::UnsupportedSource)
        ),
        "versioned external owner: unexpected result {result:?}"
    );
    assert_eq!(exact_bytes(&hostile_package)?, source);

    let ambiguous = duplicate_current_external_reference(&inserted_bytes)?;
    let ambiguous_package = Package::from_bytes(&ambiguous)?;
    let ambiguous_source = exact_bytes(&ambiguous_package)?;
    let mut ambiguous_edit = ambiguous_package.edit_body_footnote(Selector::index(0))?;
    ambiguous_edit.clear();
    let ambiguous_result = ambiguous_edit.commit();
    assert!(
        matches!(
            &ambiguous_result,
            Err(BodyFootnoteError::InvalidSource
                | BodyFootnoteError::UnsupportedDependency
                | BodyFootnoteError::UnsupportedSource)
        ),
        "ambiguous external owner: unexpected result {ambiguous_result:?}"
    );
    assert_eq!(exact_bytes(&ambiguous_package)?, ambiguous_source);
    Ok(())
}

#[test]
fn selectors_and_empty_graph_edges_are_typed_and_atomic() -> TestResult {
    let fixture = build_fixture()?;
    let package = Package::from_bytes(&fixture.bytes)?;
    let source = exact_bytes(&package)?;
    assert!(matches!(
        package.edit_body_footnote(Selector::index(99)),
        Err(BodyFootnoteError::NotFound)
    ));
    assert!(matches!(
        package.edit_body_footnote(Selector::at(Position::from_utf16_index(2)?)),
        Err(BodyFootnoteError::NotFound)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn insertion_without_view_state_only_advances_document_metadata() -> TestResult {
    let fixture = build_fixture()?;
    let bytes = without_view_state(&fixture.bytes)?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package.insert_body_footnote(Position::ZERO, "without view state", None)?;
    let target = exact_bytes(commit.package())?;
    let metadata = metadata(&target)?;
    assert_eq!(metadata.save_token, Some(ROOT_SAVE_TOKEN + 1));
    assert_eq!(
        component(&metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(DOCUMENT_SAVE_TOKEN + 1)
    );
    assert!(component(&metadata, 2).is_err());
    assert!(!has_entry(&target, VIEW_STATE_MEMBER)?);
    Ok(())
}

#[test]
fn insertion_accepts_scalar_and_story_boundaries_around_astral_text() -> TestResult {
    let fixture = build_fixture()?;
    for index in [0, 1, 4, 7] {
        let package = Package::from_bytes(&fixture.bytes)?;
        let position = Position::from_utf16_index(index)?;
        let commit = package.insert_body_footnote(position, "boundary", None)?;
        let notes = commit.package().body_footnotes()?;
        assert!(notes.iter().any(|note| note.position == position));
        assert!(commit.diagnostics().full_reparse_performed());
    }
    Ok(())
}

#[test]
fn output_limit_rejects_lifecycle_candidate_without_mutating_source() -> TestResult {
    let fixture = build_fixture()?;
    let no_previews = without_member(&fixture.bytes, PREVIEW_NAMES[0])?;
    let no_previews = without_member(&no_previews, PREVIEW_NAMES[1])?;
    let no_previews = without_member(&no_previews, PREVIEW_NAMES[2])?;
    let source_limit = u64::try_from(no_previews.len())?;
    let limits = Limits::new(
        source_limit,
        64,
        source_limit.saturating_mul(2),
        source_limit.saturating_mul(2),
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    let package = Package::from_bytes_with_limits(&no_previews, limits)?;
    let source = exact_bytes(&package)?;
    let result = package.insert_body_footnote(Position::ZERO, "limit", None);
    assert!(matches!(
        result,
        Err(BodyFootnoteError::LimitExceeded {
            kind: litchi_pages::BodyFootnoteLimitKind::InputBytes
                | litchi_pages::BodyFootnoteLimitKind::OutputBytes
                | litchi_pages::BodyFootnoteLimitKind::EntryBytes
                | litchi_pages::BodyFootnoteLimitKind::TotalBytes,
            ..
        })
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}
