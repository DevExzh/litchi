//! Integration coverage for selector-first Pages header/footer text.
//!
//! This fixture intentionally has more logical slots than physical text
//! storages: the first-section first-page header and the first-section
//! even-page footer share one storage.  A transaction must preserve that
//! aliasing, update both logical views, and leave every unrelated component
//! and metadata record byte-exact.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsk, tsp, tswp};
use litchi_pages::header_footer::{HeaderFooterSelector, Kind, Template};
use litchi_pages::{Package, Position, SectionSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const CROSS_MEMBER: &str = "Index/HeaderStorage.iwa";
const SENTINEL_MEMBER: &str = "Data/header-footer-sentinel.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const DOCUMENT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const FIRST_TEMPLATES: [u64; 3] = [60, 61, 62];
const SECOND_TEMPLATES: [u64; 3] = [63, 64, 65];
const FIRST_STORAGE_IDENTIFIER: u64 = 100;
const LAST_STORAGE_IDENTIFIER: u64 = 109;
const CROSS_STORAGE_IDENTIFIER: u64 = 700;
const VIEW_STATE_IDENTIFIER: u64 = 500;
const METADATA_IDENTIFIER: u64 = 900;
const UNRELATED_IDENTIFIER: u64 = 800;

const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const TEMPLATE_MESSAGE_TYPE: u32 = 10_143;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const LAST_OBJECT_IDENTIFIER: u64 = 1_000;
const ROOT_SAVE_TOKEN: u64 = 10;
const DOCUMENT_SAVE_TOKEN: u64 = 10;
const VIEW_STATE_SAVE_TOKEN: u64 = 7;
const UNRELATED_SAVE_TOKEN: u64 = 5;
const VERSIONED_SAVE_TOKEN: u64 = 3;
const UNKNOWN_ROOT_FIELD: u32 = 98;
const UNKNOWN_COMPONENT_FIELD: u32 = 90;
const UNKNOWN_STORAGE_FIELD: u32 = 97;

const SHARED_TEXT: &str = "Shared 😀 footer";

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

fn add_graph_metadata(object: &mut ArchiveObject, references: &[u64], fields: &[(&[u32], &[u64])]) {
    let info = &mut object.archive_info.message_infos[0];
    info.object_references.extend_from_slice(references);
    info.field_infos
        .extend(fields.iter().map(|(path, ids)| field_info(path, ids)));
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

fn storage_payload(text: &str) -> TestResult<Vec<u8>> {
    let mut payload = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Header as i32),
        text: vec![text.to_owned()],
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, UNKNOWN_STORAGE_FIELD, 0xfeed_beef)?;
    Ok(payload)
}

fn section_payload(name: &str, templates: [u64; 3]) -> TestResult<Vec<u8>> {
    let mut payload = tp::SectionArchive {
        name: Some(name.to_owned()),
        first_section_template_page: Some(reference(templates[0])),
        even_section_template_page: Some(reference(templates[1])),
        odd_section_template_page: Some(reference(templates[2])),
        inherit_previous_header_footer: Some(false),
        section_template_first_page_different: Some(true),
        section_template_even_odd_pages_different: Some(true),
        ..tp::SectionArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, 99, 0x1234_5678)?;
    Ok(payload)
}

fn template_payload(headers: &[u64], footers: &[u64]) -> Vec<u8> {
    tp::SectionTemplateArchive {
        headers: headers.iter().copied().map(reference).collect(),
        footers: footers.iter().copied().map(reference).collect(),
        ..tp::SectionTemplateArchive::default()
    }
    .encode_to_vec()
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    ids: impl IntoIterator<Item = u64>,
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: ids
            .into_iter()
            .enumerate()
            .map(|(index, id)| uuid_entry(id, u64::try_from(index + 1).unwrap()))
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, UNKNOWN_COMPONENT_FIELD, identifier + 0x9000)?;
    Ok(payload)
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let mut payload = tsp::PackageMetadata {
        last_object_identifier: LAST_OBJECT_IDENTIFIER,
        save_token: Some(ROOT_SAVE_TOKEN),
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
            FIRST_STORAGE_IDENTIFIER..=LAST_STORAGE_IDENTIFIER,
        )?,
    )?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(2, "ViewState", VIEW_STATE_SAVE_TOKEN, std::iter::empty())?,
    )?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(3, "Unrelated", UNRELATED_SAVE_TOKEN, std::iter::empty())?,
    )?;
    append_length_delimited_field(
        &mut payload,
        11,
        &metadata_component(
            DOCUMENT_IDENTIFIER,
            "Document",
            VERSIONED_SAVE_TOKEN,
            std::iter::empty(),
        )?,
    )?;
    append_varint_field(&mut payload, UNKNOWN_ROOT_FIELD, 0xdead_beef)?;
    Ok(payload)
}

fn build_fixture() -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            view_state: Some(reference(VIEW_STATE_IDENTIFIER)),
            ..tsa::DocumentArchive::default()
        },
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["Alpha\u{0004}Beta".to_owned()],
        table_section: Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(FIRST_SECTION_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 6,
                    object: Some(reference(SECOND_SECTION_IDENTIFIER)),
                },
            ],
        }),
        ..tswp::StorageArchive::default()
    };

    let mut root_object = object(
        DOCUMENT_IDENTIFIER,
        DOCUMENT_MESSAGE_TYPE,
        root.encode_to_vec(),
    )?;
    add_graph_metadata(
        &mut root_object,
        &[BODY_IDENTIFIER, VIEW_STATE_IDENTIFIER],
        &[
            (&[4], &[BODY_IDENTIFIER]),
            (&[15, 5], &[VIEW_STATE_IDENTIFIER]),
        ],
    );
    let mut body_object = object(BODY_IDENTIFIER, STORAGE_MESSAGE_TYPE, body.encode_to_vec())?;
    add_graph_metadata(
        &mut body_object,
        &[FIRST_SECTION_IDENTIFIER, SECOND_SECTION_IDENTIFIER],
        &[(
            &[17, 1, 2],
            &[FIRST_SECTION_IDENTIFIER, SECOND_SECTION_IDENTIFIER],
        )],
    );

    let mut objects = vec![root_object, body_object];
    let sections = [
        (FIRST_SECTION_IDENTIFIER, "Intro", FIRST_TEMPLATES),
        (SECOND_SECTION_IDENTIFIER, "Appendix", SECOND_TEMPLATES),
    ];
    for (identifier, name, templates) in sections {
        let mut section = object(
            identifier,
            SECTION_MESSAGE_TYPE,
            section_payload(name, templates)?,
        )?;
        add_graph_metadata(
            &mut section,
            &templates,
            &[
                (&[23], &[templates[0]]),
                (&[24], &[templates[1]]),
                (&[25], &[templates[2]]),
            ],
        );
        objects.push(section);
    }

    // Intro first header and Intro even footer intentionally share storage 100.
    let template_slots: [(u64, &[u64], &[u64]); 6] = [
        (60, &[100], &[101]),
        (61, &[102], &[100]),
        (62, &[103], &[104]),
        (63, &[105], &[106]),
        (64, &[107], &[108]),
        (65, &[109], &[100]),
    ];
    for (identifier, headers, footers) in template_slots {
        let mut template = object(
            identifier,
            TEMPLATE_MESSAGE_TYPE,
            template_payload(headers, footers),
        )?;
        let references = headers
            .iter()
            .chain(footers.iter())
            .copied()
            .collect::<Vec<_>>();
        add_graph_metadata(
            &mut template,
            &references,
            &[(&[1], headers), (&[2], footers)],
        );
        objects.push(template);
    }

    for identifier in FIRST_STORAGE_IDENTIFIER..=LAST_STORAGE_IDENTIFIER {
        let text = match identifier {
            100 => SHARED_TEXT,
            101 => "Intro first footer",
            102 => "Intro even header",
            103 => "Intro odd header",
            104 => "Intro odd footer",
            105 => "Appendix first header",
            106 => "Appendix first footer",
            107 => "Appendix even header",
            108 => "Appendix even footer",
            109 => "Appendix odd header",
            _ => unreachable!(),
        };
        let storage = object(identifier, STORAGE_MESSAGE_TYPE, storage_payload(text)?)?;
        objects.push(storage);
    }

    let document_component = compressed(objects)?;
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
    let sentinel = b"sentinel-preservation-witness".to_vec();
    let entries = [
        (DOCUMENT_MEMBER, document_component.as_slice()),
        (METADATA_MEMBER, metadata_component.as_slice()),
        (VIEW_STATE_MEMBER, view_state_component.as_slice()),
        (UNRELATED_MEMBER, unrelated_component.as_slice()),
        (SENTINEL_MEMBER, sentinel.as_slice()),
        (PREVIEWS[0], b"preview-full".as_slice()),
        (PREVIEWS[1], b"preview-micro".as_slice()),
        (PREVIEWS[2], b"preview-web".as_slice()),
    ];
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
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
        .ok_or_else(|| io::Error::other("metadata member is missing"))?;
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

fn metadata_component_raw(bytes: &[u8], identifier: u64) -> TestResult<Vec<Vec<u8>>> {
    let payload = metadata_payload_from_bytes(bytes)?;
    let view = WireView::parse(&payload)?;
    Ok(view
        .fields()
        .filter(|field| field.number() == 3)
        .filter_map(|field| {
            tsp::ComponentInfo::decode(field.payload())
                .ok()
                .filter(|component| component.identifier == identifier)
                .map(|_| field.payload().to_vec())
        })
        .collect())
}

fn component(metadata: &tsp::PackageMetadata, identifier: u64) -> TestResult<&tsp::ComponentInfo> {
    metadata
        .components
        .iter()
        .find(|component| component.identifier == identifier)
        .ok_or_else(|| io::Error::other("metadata component is missing").into())
}

fn member_bytes(bytes: &[u8], member: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(bytes)?
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("member is missing"))?
        .data()
        .to_vec())
}

fn object_payload(bytes: &[u8], member: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("member is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    archive
        .object(identifier)
        .and_then(|object| object.messages.first())
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("object message is missing").into())
}

fn has_member(bytes: &[u8], member: &str) -> TestResult<bool> {
    Ok(Catalog::from_bytes(bytes)?
        .iter()
        .any(|entry| entry.name() == member))
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
        .ok_or_else(|| io::Error::other("member is missing"))?;
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

fn remove_member(source: &[u8], member: &str) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
            &[],
            &[member],
            Limits::default(),
        )?,
    )
}

fn selector(section: usize, template: Template, kind: Kind) -> HeaderFooterSelector<'static> {
    HeaderFooterSelector::index(section, template, kind, 0)
}

fn assert_alias_text(package: &Package, expected: &str) -> TestResult {
    assert_eq!(
        package
            .edit_header_footer_text(selector(0, Template::First, Kind::Header))?
            .text(),
        expected
    );
    assert_eq!(
        package
            .edit_header_footer_text(selector(0, Template::Even, Kind::Footer))?
            .text(),
        expected
    );
    assert_eq!(
        package
            .edit_header_footer_text(selector(1, Template::Odd, Kind::Footer))?
            .text(),
        expected
    );
    Ok(())
}

fn assert_unknown_storage_fields(bytes: &[u8], identifier: u64) -> TestResult {
    let catalog = Catalog::from_bytes(bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("document member is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(entry.data())?.into_bytes())?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("storage object is missing"))?;
    let payload = &object
        .messages
        .first()
        .ok_or_else(|| io::Error::other("storage message is missing"))?
        .data;
    let view = WireView::parse(payload)?;
    assert!(
        view.fields()
            .any(|field| field.number() == UNKNOWN_STORAGE_FIELD)
    );
    Ok(())
}

fn assert_metadata_unknown_fields(bytes: &[u8]) -> TestResult {
    let payload = metadata_payload_from_bytes(bytes)?;
    let view = WireView::parse(&payload)?;
    assert!(
        view.fields()
            .any(|field| field.number() == UNKNOWN_ROOT_FIELD)
    );
    for field in view.fields().filter(|field| field.number() == 3) {
        let nested = WireView::parse(field.payload())?;
        assert!(
            nested
                .fields()
                .any(|nested_field| nested_field.number() == UNKNOWN_COMPONENT_FIELD)
        );
    }
    Ok(())
}

fn remove_storage(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(identifier));
        Ok(())
    })
}

fn duplicate_storage_kind(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_STORAGE_IDENTIFIER,
        |_type_, mut data| {
            append_varint_field(&mut data, 1, 1)?;
            Ok((STORAGE_MESSAGE_TYPE, data))
        },
    )
}

fn wrong_storage_kind_wire(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_STORAGE_IDENTIFIER,
        |_type_, mut data| {
            append_length_delimited_field(&mut data, 1, &[0])?;
            Ok((STORAGE_MESSAGE_TYPE, data))
        },
    )
}

fn valid_non_header_storage_kind(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_STORAGE_IDENTIFIER,
        |_type_, data| {
            let mut storage = tswp::StorageArchive::decode(data.as_slice())?;
            storage.kind = Some(tswp::storage_archive::KindType::Body as i32);
            Ok((STORAGE_MESSAGE_TYPE, storage.encode_to_vec()))
        },
    )
}

fn add_unattributed_field_owner(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, UNRELATED_MEMBER, |archive| {
        let object = archive
            .object_mut(UNRELATED_IDENTIFIER)
            .ok_or_else(|| io::Error::other("unrelated object is missing"))?;
        object.archive_info.message_infos[0]
            .field_infos
            .push(field_info(&[99], &[FIRST_STORAGE_IDENTIFIER]));
        Ok(())
    })
}

fn wrong_template_reference_path(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(FIRST_TEMPLATES[0])
            .ok_or_else(|| io::Error::other("template object is missing"))?;
        let field = object.archive_info.message_infos[0]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [1])
            .ok_or_else(|| io::Error::other("header FieldInfo is missing"))?;
        field.path = vec![99].into();
        Ok(())
    })
}

#[derive(Clone, Copy)]
enum MetadataOwnerKind {
    ForeignUuid,
    External,
    Data,
    Ambiguous,
    RootDataMap,
}

fn add_metadata_storage_owner(source: &[u8], kind: MetadataOwnerKind) -> TestResult<Vec<u8>> {
    let mut metadata = metadata(source)?;
    let document = metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == DOCUMENT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("Document metadata component is missing"))?;
    match kind {
        MetadataOwnerKind::ForeignUuid => {
            let unrelated = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == 3)
                .ok_or_else(|| io::Error::other("Unrelated metadata component is missing"))?;
            unrelated
                .object_uuid_map_entries
                .push(uuid_entry(FIRST_STORAGE_IDENTIFIER, 0xfeed));
        },
        MetadataOwnerKind::External => {
            document
                .external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: DOCUMENT_IDENTIFIER,
                    object_identifier: Some(FIRST_STORAGE_IDENTIFIER),
                    is_weak: None,
                });
        },
        MetadataOwnerKind::Data => {
            document.data_references.push(tsp::ComponentDataReference {
                data_identifier: 77,
                object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                    object_identifier: FIRST_STORAGE_IDENTIFIER,
                    count: 1,
                }],
            });
        },
        MetadataOwnerKind::Ambiguous => {
            document
                .ambiguous_object_identifiers
                .push(FIRST_STORAGE_IDENTIFIER);
        },
        MetadataOwnerKind::RootDataMap => {
            metadata.data_metadata_map = Some(reference(FIRST_STORAGE_IDENTIFIER));
        },
    }
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, metadata.encode_to_vec())),
    )
}

fn duplicate_root_body_identifier(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("document member is missing"))?;
    let archive = Archive::parse(&SnappyStream::decompress(document.data())?.into_bytes())?;
    let duplicate = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(BODY_IDENTIFIER))
        .cloned()
        .ok_or_else(|| io::Error::other("body object is missing"))?;
    let duplicate = compressed(vec![duplicate])?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push(("Index/DuplicateBody.iwa", duplicate.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn ambiguous_section_name(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        SECOND_SECTION_IDENTIFIER,
        |_type_, data| {
            let mut section = tp::SectionArchive::decode(data.as_slice())?;
            section.name = Some("Intro".to_owned());
            Ok((SECTION_MESSAGE_TYPE, section.encode_to_vec()))
        },
    )
}

fn duplicate_metadata_component(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_from_bytes(source)?;
    append_length_delimited_field(
        &mut payload,
        3,
        &metadata_component(
            DOCUMENT_IDENTIFIER,
            "Document",
            DOCUMENT_SAVE_TOKEN,
            std::iter::empty(),
        )?,
    )?;
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn duplicate_metadata_save_token(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_from_bytes(source)?;
    append_varint_field(&mut payload, 8, ROOT_SAVE_TOKEN + 1)?;
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn wrong_metadata_save_token_wire(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_from_bytes(source)?;
    append_length_delimited_field(&mut payload, 8, &[1])?;
    rewrite_object_payload(
        source,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, payload)),
    )
}

fn cross_component_storage(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut changed = rewrite_object_payload(
        source,
        DOCUMENT_MEMBER,
        FIRST_TEMPLATES[0],
        |_type_, data| {
            let mut template = tp::SectionTemplateArchive::decode(data.as_slice())?;
            template.headers[0] = reference(CROSS_STORAGE_IDENTIFIER);
            Ok((TEMPLATE_MESSAGE_TYPE, template.encode_to_vec()))
        },
    )?;
    changed = rewrite_member(&changed, DOCUMENT_MEMBER, |archive| {
        let template = archive
            .object_mut(FIRST_TEMPLATES[0])
            .ok_or_else(|| io::Error::other("template object is missing"))?;
        let info = template
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("template message info is missing"))?;
        for identifier in &mut info.object_references {
            if *identifier == FIRST_STORAGE_IDENTIFIER {
                *identifier = CROSS_STORAGE_IDENTIFIER;
            }
        }
        for field in &mut info.field_infos {
            if field.path.as_slice() == [1] {
                for identifier in &mut field.object_references {
                    if *identifier == FIRST_STORAGE_IDENTIFIER {
                        *identifier = CROSS_STORAGE_IDENTIFIER;
                    }
                }
            }
        }
        Ok(())
    })?;
    // The foreign storage is a real current component, not an unregistered
    // object hidden behind the Document template.  Its exact selector is
    // what lets the package owner advance only this component's token.
    let mut metadata = metadata_payload_from_bytes(&changed)?;
    append_length_delimited_field(
        &mut metadata,
        3,
        &metadata_component(
            4,
            "HeaderStorage",
            8,
            std::iter::once(CROSS_STORAGE_IDENTIFIER),
        )?,
    )?;
    changed = rewrite_object_payload(
        &changed,
        METADATA_MEMBER,
        METADATA_IDENTIFIER,
        |_type_, _data| Ok((METADATA_MESSAGE_TYPE, metadata)),
    )?;
    let cross = compressed(vec![object(
        CROSS_STORAGE_IDENTIFIER,
        STORAGE_MESSAGE_TYPE,
        storage_payload(SHARED_TEXT)?,
    )?])?;
    let catalog = Catalog::from_bytes(&changed)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((CROSS_MEMBER, cross.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn without_metadata(source: &[u8]) -> TestResult<Vec<u8>> {
    remove_member(source, METADATA_MEMBER)
}

#[test]
fn enumeration_selectors_and_aliases_are_semantic() -> TestResult {
    let source = build_fixture()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.header_footers()?.len(), 12);
    assert_eq!(
        package
            .edit_header_footer_text(selector(0, Template::First, Kind::Header))?
            .text(),
        SHARED_TEXT
    );
    assert_alias_text(&package, SHARED_TEXT)?;
    assert_eq!(
        package
            .edit_header_footer_text(HeaderFooterSelector::new(
                SectionSelector::name("Intro"),
                Template::First,
                Kind::Header,
                Position::new(0),
            ))?
            .text(),
        SHARED_TEXT
    );
    assert!(
        package
            .edit_header_footer_text(selector(8, Template::Odd, Kind::Header))
            .is_err()
    );
    Ok(())
}

#[test]
fn set_preserves_aliases_metadata_and_unknown_bytes_and_inverts_exactly() -> TestResult {
    let source = build_fixture()?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_header_footer_text(selector(0, Template::First, Kind::Header))?;
    edit.set("Updated header/footer — 東京")?;
    let commit = edit.commit()?;
    let target = exact_bytes(commit.package())?;
    assert_ne!(target, source);
    assert_alias_text(commit.package(), "Updated header/footer — 東京")?;
    assert!(commit.diagnostics().changed_bytes());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(
        member_bytes(&target, SENTINEL_MEMBER)?,
        member_bytes(&source, SENTINEL_MEMBER)?
    );
    assert_eq!(
        member_bytes(&target, UNRELATED_MEMBER)?,
        member_bytes(&source, UNRELATED_MEMBER)?
    );
    assert_eq!(
        object_payload(&target, DOCUMENT_MEMBER, 101)?,
        object_payload(&source, DOCUMENT_MEMBER, 101)?
    );
    assert!(!has_member(&target, PREVIEWS[0])?);
    assert!(!has_member(&target, PREVIEWS[1])?);
    assert!(!has_member(&target, PREVIEWS[2])?);
    assert_unknown_storage_fields(&target, FIRST_STORAGE_IDENTIFIER)?;
    assert_metadata_unknown_fields(&target)?;

    let source_metadata = metadata(&source)?;
    let target_metadata = metadata(&target)?;
    assert_eq!(
        target_metadata.last_object_identifier,
        source_metadata.last_object_identifier
    );
    assert_eq!(target_metadata.save_token, Some(ROOT_SAVE_TOKEN + 1));
    assert_eq!(
        component(&target_metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(DOCUMENT_SAVE_TOKEN + 1)
    );
    assert_eq!(
        component(&target_metadata, 2)?.save_token,
        Some(VIEW_STATE_SAVE_TOKEN)
    );
    assert_eq!(
        component(&target_metadata, 3)?.save_token,
        Some(UNRELATED_SAVE_TOKEN)
    );
    assert_eq!(
        metadata_component_raw(&target, 2)?,
        metadata_component_raw(&source, 2)?
    );
    assert_eq!(
        metadata_component_raw(&target, 3)?,
        metadata_component_raw(&source, 3)?
    );
    assert_eq!(
        target_metadata.versioned_components[0].save_token,
        Some(VERSIONED_SAVE_TOKEN)
    );
    let restored = commit
        .package()
        .apply_header_footer_text(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    let forward = package.apply_header_footer_text(commit.patch())?;
    assert_eq!(exact_bytes(forward.package())?, target);
    Ok(())
}

#[test]
fn utf16_replace_and_clear_are_alias_aware() -> TestResult {
    let source = build_fixture()?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_header_footer_text(selector(0, Template::Even, Kind::Footer))?;
    edit.replace(7..9, "北")?;
    let replaced = edit.commit()?;
    assert_alias_text(replaced.package(), "Shared 北 footer")?;

    let mut clear =
        replaced
            .package()
            .edit_header_footer_text(selector(0, Template::First, Kind::Header))?;
    clear.clear()?;
    let cleared = clear.commit()?;
    assert_alias_text(cleared.package(), "")?;
    let restored = cleared
        .package()
        .apply_header_footer_text(&cleared.patch().inverse())?;
    assert_eq!(
        exact_bytes(restored.package())?,
        exact_bytes(replaced.package())?
    );
    Ok(())
}

#[test]
fn no_op_and_patch_conflict_are_exact_and_atomic() -> TestResult {
    let source = build_fixture()?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_header_footer_text(selector(0, Template::First, Kind::Header))?;
    edit.set(SHARED_TEXT)?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed_bytes());
    assert_eq!(exact_bytes(noop.package())?, source);

    let mut edit = package.edit_header_footer_text(selector(0, Template::First, Kind::Header))?;
    edit.set("conflict")?;
    let changed = edit.commit()?;
    let tampered_source = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[EntryEdit::new(SENTINEL_MEMBER, b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(tampered.apply_header_footer_text(changed.patch()).is_err());
    assert_eq!(exact_bytes(&tampered)?, tampered_source);
    Ok(())
}

#[test]
fn malformed_missing_and_ambiguous_slots_fail_closed() -> TestResult {
    let source = build_fixture()?;
    for malformed in [
        duplicate_storage_kind(&source)?,
        wrong_storage_kind_wire(&source)?,
        valid_non_header_storage_kind(&source)?,
        remove_storage(&source, FIRST_STORAGE_IDENTIFIER)?,
        add_unattributed_field_owner(&source)?,
        wrong_template_reference_path(&source)?,
        duplicate_root_body_identifier(&source)?,
    ] {
        if let Ok(package) = Package::from_bytes(&malformed) {
            let before = exact_bytes(&package)?;
            let result = package
                .edit_header_footer_text(selector(0, Template::First, Kind::Header))
                .and_then(|mut edit| {
                    edit.set("rejected")?;
                    edit.commit()
                });
            assert!(result.is_err());
            assert_eq!(exact_bytes(&package)?, before);
        }
    }
    let ambiguous = ambiguous_section_name(&source)?;
    if let Ok(package) = Package::from_bytes(&ambiguous) {
        let before = exact_bytes(&package)?;
        let result = package
            .edit_header_footer_text(HeaderFooterSelector::new(
                SectionSelector::name("Intro"),
                Template::First,
                Kind::Header,
                Position::new(0),
            ))
            .and_then(|mut edit| {
                edit.set("rejected")?;
                edit.commit()
            });
        assert!(result.is_err());
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn metadata_storage_ownership_routes_fail_closed() -> TestResult {
    let source = build_fixture()?;
    for kind in [
        MetadataOwnerKind::ForeignUuid,
        MetadataOwnerKind::External,
        MetadataOwnerKind::Data,
        MetadataOwnerKind::Ambiguous,
        MetadataOwnerKind::RootDataMap,
    ] {
        let hostile = add_metadata_storage_owner(&source, kind)?;
        let package = Package::from_bytes(&hostile)?;
        let before = exact_bytes(&package)?;
        let result = package
            .edit_header_footer_text(selector(0, Template::First, Kind::Header))
            .and_then(|mut edit| {
                edit.set("rejected")?;
                edit.commit()
            });
        assert!(result.is_err());
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn cross_component_storage_is_a_supported_exact_owner() -> TestResult {
    let source = cross_component_storage(&build_fixture()?)?;
    let package = Package::from_bytes(&source)?;
    let document_before = member_bytes(&source, DOCUMENT_MEMBER)?;
    let cross_before = member_bytes(&source, CROSS_MEMBER)?;
    let metadata_before = member_bytes(&source, METADATA_MEMBER)?;
    let selected = selector(0, Template::First, Kind::Header);

    assert_eq!(
        package.edit_header_footer_text(selected)?.text(),
        SHARED_TEXT
    );
    // The other two logical aliases still point at the original Document
    // storage; only the selected slot is routed to the foreign component.
    assert_eq!(
        package
            .edit_header_footer_text(selector(0, Template::Even, Kind::Footer))?
            .text(),
        SHARED_TEXT
    );
    assert_eq!(
        package
            .edit_header_footer_text(selector(1, Template::Odd, Kind::Footer))?
            .text(),
        SHARED_TEXT
    );

    let mut edit = package.edit_header_footer_text(selected)?;
    edit.set("Cross component — 北区")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().edit_header_footer_text(selected)?.text(),
        "Cross component — 北区"
    );
    assert_eq!(
        commit
            .package()
            .edit_header_footer_text(selector(0, Template::Even, Kind::Footer))?
            .text(),
        SHARED_TEXT
    );
    assert_eq!(
        commit
            .package()
            .edit_header_footer_text(selector(1, Template::Odd, Kind::Footer))?
            .text(),
        SHARED_TEXT
    );
    let target = exact_bytes(commit.package())?;
    assert_eq!(member_bytes(&target, DOCUMENT_MEMBER)?, document_before);
    assert_ne!(member_bytes(&target, CROSS_MEMBER)?, cross_before);
    assert_ne!(member_bytes(&target, METADATA_MEMBER)?, metadata_before);
    assert_eq!(
        member_bytes(&target, VIEW_STATE_MEMBER)?,
        member_bytes(&source, VIEW_STATE_MEMBER)?
    );
    assert_eq!(
        member_bytes(&target, UNRELATED_MEMBER)?,
        member_bytes(&source, UNRELATED_MEMBER)?
    );
    assert_eq!(
        member_bytes(&target, SENTINEL_MEMBER)?,
        member_bytes(&source, SENTINEL_MEMBER)?
    );
    assert_eq!(
        metadata(&target)?.last_object_identifier,
        LAST_OBJECT_IDENTIFIER
    );

    let source_metadata = metadata(&source)?;
    let target_metadata = metadata(&target)?;
    assert_eq!(source_metadata.save_token, Some(ROOT_SAVE_TOKEN));
    assert_eq!(target_metadata.save_token, Some(ROOT_SAVE_TOKEN + 1));
    assert_eq!(
        component(&source_metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(DOCUMENT_SAVE_TOKEN)
    );
    assert_eq!(
        component(&target_metadata, DOCUMENT_IDENTIFIER)?.save_token,
        Some(DOCUMENT_SAVE_TOKEN)
    );
    assert_eq!(component(&source_metadata, 4)?.save_token, Some(8));
    assert_eq!(
        component(&target_metadata, 4)?.save_token,
        Some(ROOT_SAVE_TOKEN + 1)
    );
    assert_eq!(
        component(&target_metadata, 2)?.save_token,
        Some(VIEW_STATE_SAVE_TOKEN)
    );
    assert_eq!(
        component(&target_metadata, 3)?.save_token,
        Some(UNRELATED_SAVE_TOKEN)
    );
    assert!(!has_member(&target, PREVIEWS[0])?);
    assert!(!has_member(&target, PREVIEWS[1])?);
    assert!(!has_member(&target, PREVIEWS[2])?);
    assert_eq!(commit.diagnostics().touched_components(), 2);

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_header_footer_text(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn metadata_identity_and_missing_metadata_fail_closed() -> TestResult {
    let source = build_fixture()?;
    for malformed in [
        duplicate_metadata_component(&source)?,
        duplicate_metadata_save_token(&source)?,
        wrong_metadata_save_token_wire(&source)?,
        without_metadata(&source)?,
    ] {
        if let Ok(package) = Package::from_bytes(&malformed) {
            let before = exact_bytes(&package)?;
            let result = package
                .edit_header_footer_text(selector(0, Template::First, Kind::Header))
                .and_then(|mut edit| {
                    edit.set("rejected")?;
                    edit.commit()
                });
            assert!(result.is_err());
            assert_eq!(exact_bytes(&package)?, before);
        }
    }
    Ok(())
}

#[test]
fn archive_limits_fail_before_publication() -> TestResult {
    let source = build_fixture()?;
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let before = exact_bytes(&package)?;
        let result = package
            .edit_header_footer_text(selector(0, Template::First, Kind::Header))
            .and_then(|mut edit| {
                edit.set("bounded")?;
                edit.commit()
            });
        assert!(result.is_err());
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn output_ceiling_rejects_growth_after_ingress_and_preserves_source() -> TestResult {
    let source = build_fixture()?;
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(source.len())?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let package = Package::from_bytes_with_limits(&source, limits)?;
    let before = exact_bytes(&package)?;
    let mut state = 0x9e37_79b9_7f4a_7c15u64;
    let replacement: String = (0..65_536)
        .map(|_| {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            char::from(b'a' + ((state >> 32) % 26) as u8)
        })
        .collect();
    let result = package
        .edit_header_footer_text(selector(0, Template::First, Kind::Header))
        .and_then(|mut edit| {
            edit.set(&replacement)?;
            edit.commit()
        });
    assert!(result.is_err());
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn physical_entry_stream_total_and_nesting_limits_fail_closed() -> TestResult {
    let source = build_fixture()?;
    let defaults = Limits::default();
    let profiles = [
        Limits::new(
            defaults.max_input_bytes(),
            defaults.max_entries(),
            1,
            defaults.max_total_bytes(),
            defaults.max_iwa_stream_bytes(),
        )?,
        Limits::new(
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            1,
            defaults.max_iwa_stream_bytes(),
        )?,
        Limits::new(
            defaults.max_input_bytes(),
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_total_bytes(),
            1,
        )?,
    ];
    for limits in profiles {
        assert!(Package::from_bytes_with_limits(&source, limits).is_err());
    }
    let archive_limits = litchi_iwa_core::Limits::default().with_header_nesting(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    assert!(Package::from_bytes_with_limits(&source, limits).is_err());
    Ok(())
}

#[test]
fn public_transaction_types_are_send_sync_and_debug_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<HeaderFooterSelector<'static>>();
    assert_send_sync_debug::<litchi_pages::HeaderFooterTextEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::HeaderFooterTextPatch>();
    assert_send_sync_debug::<litchi_pages::HeaderFooterTextCommit>();
    assert_send_sync_debug::<litchi_pages::HeaderFooterTextDiagnostics>();
    let package = Package::from_bytes(&build_fixture()?)?;
    let edit = package.edit_header_footer_text(selector(0, Template::First, Kind::Header))?;
    let debug = format!("{edit:?}");
    assert!(!debug.contains(DOCUMENT_MEMBER));
    Ok(())
}
