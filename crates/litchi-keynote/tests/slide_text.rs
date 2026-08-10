use std::io;
use std::path::PathBuf;

use litchi_iwa_archive::{Limits, package::Catalog, package::EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp, tswp};
use litchi_keynote::{
    Package, Position, ReadOptions, SemanticLimits, SlideSelector, SlideTextCommit,
    SlideTextDiagnostics, SlideTextEdit, SlideTextError, SlideTextLimitKind, SlideTextPatch,
    SlideTextRole, TextPosition, TextSpan,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const PREVIEW_MEMBERS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const FIRST_SLIDE: u64 = 4;
const TITLE_OWNER: u64 = 5;
const BODY_OWNER: u64 = 6;
const SIBLING_OWNER: u64 = 7;
const TITLE_STORAGE: u64 = 10;
const BODY_STORAGE: u64 = 11;
const SIBLING_STORAGE: u64 = 12;
const SECOND_SLIDE: u64 = 31;
const SECOND_TITLE_OWNER: u64 = 32;
const SECOND_TITLE_STORAGE: u64 = 33;
const UNRELATED_PLACEHOLDER: u64 = 501;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const SHAPE_MESSAGE_TYPE: u32 = 2_011;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const PRIVATE_MARKER: &[u8] = b"private-keynote-slide-text-marker-2147483647";
const UNRELATED_PLACEHOLDER_MARKER: &[u8] = b"unrelated-hidden-placeholder-marker";

const TITLE: &str = "Launch 🚀 title";
const BODY: &str = "Body 東京😀";
const SIBLING: &str = "2026-08-09 sibling text";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, value: &impl prost::Message) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_,
            data: value.encode_to_vec(),
        }],
    )?)
}

fn storage(identifier: u64, fragments: &[&str]) -> TestResult<ArchiveObject> {
    let mut payload = tswp::StorageArchive {
        text: fragments
            .iter()
            .map(|fragment| (*fragment).to_owned())
            .collect(),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut payload, 98, identifier)?;
    Ok(ArchiveObject::new(
        identifier,
        vec![
            RawMessage {
                type_: 779,
                data: format!("before-storage-{identifier}").into_bytes(),
            },
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: payload,
            },
            RawMessage {
                type_: 780,
                data: format!("after-storage-{identifier}").into_bytes(),
            },
        ],
    )?)
}

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let bytes = Archive { objects }.to_bytes()?;
    let parsed = Archive::parse(&bytes)?;
    let object = parsed
        .object(target_identifier)
        .ok_or_else(|| io::Error::other("synthetic target object is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length_u64, prefix_length) = decode_varint_from_bytes(
        bytes
            .get(header_offset..)
            .ok_or_else(|| io::Error::other("synthetic header offset is invalid"))?,
    )?;
    let header_length = usize::try_from(header_length_u64)?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    if header_end != data_offset {
        return Err(io::Error::other("synthetic object offsets disagree").into());
    }

    let mut header = bytes
        .get(header_start..header_end)
        .ok_or_else(|| io::Error::other("synthetic header range is invalid"))?
        .to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut with_unknown = Vec::with_capacity(bytes.len().saturating_add(8));
    with_unknown.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut with_unknown, u64::try_from(header.len())?);
    with_unknown.extend_from_slice(&header);
    with_unknown.extend_from_slice(&bytes[data_offset..]);

    let reparsed = Archive::parse(&with_unknown)?;
    assert_eq!(reparsed.to_bytes()?, with_unknown);
    Ok(SnappyStream::compress(&with_unknown)?)
}

fn slide(
    identifier: u64,
    name: &str,
    title: Option<u64>,
    body: Option<u64>,
    drawables: Vec<u64>,
) -> kn::SlideArchive {
    kn::SlideArchive {
        style: reference(identifier.saturating_add(1_000)),
        transition: kn::TransitionArchive::default(),
        title_placeholder: title.map(reference),
        body_placeholder: body.map(reference),
        owned_drawables: drawables.into_iter().map(reference).collect(),
        name: Some(name.to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    }
}

fn placeholder(storage_identifier: u64) -> kn::PlaceholderArchive {
    placeholder_storage_topology(Some(storage_identifier), None)
}

#[allow(
    deprecated,
    reason = "the adversarial fixture exercises native Keynote's legacy storage edge"
)]
fn placeholder_storage_topology(
    owned_storage: Option<u64>,
    deprecated_storage: Option<u64>,
) -> kn::PlaceholderArchive {
    kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            deprecated_storage: deprecated_storage.map(reference),
            owned_storage: owned_storage.map(reference),
            ..tswp::ShapeInfoArchive::default()
        },
        ..kn::PlaceholderArchive::default()
    }
}

fn synthetic_package(dependent_title_content: bool) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(3), reference(30)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema requires legacy cache fields")]
    let first_node = kn::SlideNodeArchive {
        slide: Some(reference(FIRST_SLIDE)),
        thumbnails: vec![
            tsp::DataReference { identifier: 7_001 },
            tsp::DataReference { identifier: 7_002 },
        ],
        thumbnail_sizes: vec![
            tsp::Size {
                width: 320.0,
                height: 240.0,
            },
            tsp::Size {
                width: 160.0,
                height: 120.0,
            },
        ],
        thumbnails_are_dirty: Some(false),
        digests_for_datas_needing_download_for_thumbnail: vec!["stale-digest".to_owned()],
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    #[allow(deprecated, reason = "native schema requires legacy cache fields")]
    let second_node = kn::SlideNodeArchive {
        slide: Some(reference(SECOND_SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };

    let mut title_storage = tswp::StorageArchive {
        text: vec!["Launch ".to_owned(), "🚀 title".to_owned()],
        ..tswp::StorageArchive::default()
    };
    if dependent_title_content {
        title_storage.table_attachment = Some(tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 7,
                object: Some(reference(500)),
            }],
        });
    }
    let mut title_payload = title_storage.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut title_payload, 98, 98_998)?;

    let mut first_slide = object(
        FIRST_SLIDE,
        SLIDE_MESSAGE_TYPE,
        &slide(
            FIRST_SLIDE,
            "Agenda",
            Some(TITLE_OWNER),
            Some(BODY_OWNER),
            vec![TITLE_OWNER, BODY_OWNER, SIBLING_OWNER],
        ),
    )?;
    first_slide.archive_info.message_infos[0].object_references =
        vec![TITLE_OWNER, BODY_OWNER, SIBLING_OWNER];

    let mut title_owner = object(
        TITLE_OWNER,
        PLACEHOLDER_MESSAGE_TYPE,
        &placeholder(TITLE_STORAGE),
    )?;
    title_owner.archive_info.message_infos[0].object_references = vec![TITLE_STORAGE];
    let mut body_owner = object(
        BODY_OWNER,
        PLACEHOLDER_MESSAGE_TYPE,
        &placeholder(BODY_STORAGE),
    )?;
    body_owner.archive_info.message_infos[0].object_references = vec![BODY_STORAGE];
    let mut sibling_owner = object(
        SIBLING_OWNER,
        SHAPE_MESSAGE_TYPE,
        &tswp::ShapeInfoArchive {
            owned_storage: Some(reference(SIBLING_STORAGE)),
            ..tswp::ShapeInfoArchive::default()
        },
    )?;
    sibling_owner.archive_info.message_infos[0].object_references = vec![SIBLING_STORAGE];

    let mut title_storage_object = ArchiveObject::new(
        TITLE_STORAGE,
        vec![
            RawMessage {
                type_: 779,
                data: b"before-title-storage".to_vec(),
            },
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: title_payload,
            },
            RawMessage {
                type_: 780,
                data: b"after-title-storage".to_vec(),
            },
        ],
    )?;
    title_storage_object.archive_info.message_infos[1].object_references =
        dependent_title_content.then_some(500).into_iter().collect();

    let mut second_slide = object(
        SECOND_SLIDE,
        SLIDE_MESSAGE_TYPE,
        &slide(
            SECOND_SLIDE,
            "Existing Empty",
            Some(SECOND_TITLE_OWNER),
            None,
            vec![SECOND_TITLE_OWNER],
        ),
    )?;
    second_slide.archive_info.message_infos[0].object_references = vec![SECOND_TITLE_OWNER];
    let mut second_title_owner = object(
        SECOND_TITLE_OWNER,
        PLACEHOLDER_MESSAGE_TYPE,
        &placeholder(SECOND_TITLE_STORAGE),
    )?;
    second_title_owner.archive_info.message_infos[0].object_references = vec![SECOND_TITLE_STORAGE];

    let mut first_node_object = object(3, 4, &first_node)?;
    first_node_object.archive_info.message_infos[0].data_references = vec![7_001, 7_002];
    let mut thumbnail_field = FieldInfo::new(vec![16]);
    thumbnail_field.data_references = vec![7_001, 7_002];
    first_node_object.archive_info.message_infos[0]
        .field_infos
        .push(thumbnail_field);

    let document_component = component_with_unknown_header(
        vec![
            object(1, 1, &document)?,
            object(2, 2, &show)?,
            first_node_object,
            first_slide,
            title_owner,
            body_owner,
            sibling_owner,
            title_storage_object,
            storage(BODY_STORAGE, &["Body ", "東京😀"])?,
            storage(SIBLING_STORAGE, &[SIBLING])?,
            object(30, 4, &second_node)?,
            second_slide,
            second_title_owner,
            storage(SECOND_TITLE_STORAGE, &[""])?,
            ArchiveObject::new(
                500,
                vec![RawMessage {
                    type_: 999,
                    data: PRIVATE_MARKER.to_vec(),
                }],
            )?,
        ],
        TITLE_STORAGE,
    )?;
    let unrelated_component = SnappyStream::compress(
        &Archive {
            objects: vec![ArchiveObject::new(
                900,
                vec![RawMessage {
                    type_: 999,
                    data: b"unrelated-iwa-component".to_vec(),
                }],
            )?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            ("Index/Unrelated.iwa", unrelated_component.as_slice()),
            (PREVIEW_MEMBERS[0], b"synthetic full preview".as_slice()),
            (PREVIEW_MEMBERS[1], b"synthetic micro preview".as_slice()),
            (PREVIEW_MEMBERS[2], b"synthetic web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_document(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn with_overlong_object_length_prefix(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let offset = usize::try_from(object.header_offset)?;
    let (_length, prefix_bytes) = decode_varint_from_bytes(&stream[offset..])?;
    if prefix_bytes != 1 {
        return Err(io::Error::other("synthetic prefix is not one byte").into());
    }
    stream[offset] |= 0x80;
    stream.insert(offset + 1, 0);
    Archive::parse(&stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn with_title_storage_topology(
    package: &[u8],
    owned_storage: Option<u64>,
    deprecated_storage: Option<u64>,
) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let object = archive
            .object_mut(TITLE_OWNER)
            .ok_or_else(|| io::Error::other("missing title placeholder"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == PLACEHOLDER_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing title-placeholder message"))?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data: placeholder_storage_topology(owned_storage, deprecated_storage)
                    .encode_to_vec(),
            },
        )?;
        let mut references = Vec::with_capacity(2);
        if let Some(identifier) = owned_storage {
            references.push(identifier);
        }
        if let Some(identifier) = deprecated_storage
            && !references.contains(&identifier)
        {
            references.push(identifier);
        }
        object.archive_info.message_infos[index].object_references = references;
        Ok(())
    })
}

fn with_placeholder_kind(package: &[u8], owner: u64, kind: Option<i32>) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let object = archive
            .object_mut(owner)
            .ok_or_else(|| io::Error::other("missing selected placeholder"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == PLACEHOLDER_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing placeholder message"))?;
        let mut value = kn::PlaceholderArchive::decode(object.messages[index].data.as_slice())?;
        value.kind = kind;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data: value.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn with_owned_storage_reference_unknown(package: &[u8], owner: u64) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let object = archive
            .object_mut(owner)
            .ok_or_else(|| io::Error::other("missing selected placeholder"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == PLACEHOLDER_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing placeholder message"))?;
        let source = object.messages[index].data.as_slice();
        let placeholder = WireView::parse(source)?;
        let mut rewritten_placeholder = Vec::with_capacity(source.len() + 8);
        let mut selected_super = false;
        for field in placeholder.fields() {
            if field.number() != 1 {
                rewritten_placeholder.extend_from_slice(field.raw());
                continue;
            }
            if std::mem::replace(&mut selected_super, true) || field.wire_type() != 2 {
                return Err(io::Error::other("ambiguous placeholder super field").into());
            }
            let source_shape = field.canonical_payload()?;
            let shape = WireView::parse(source_shape)?;
            let mut rewritten_shape = Vec::with_capacity(source_shape.len() + 8);
            let mut selected_storage = false;
            for shape_field in shape.fields() {
                if shape_field.number() != 4 {
                    rewritten_shape.extend_from_slice(shape_field.raw());
                    continue;
                }
                if std::mem::replace(&mut selected_storage, true) || shape_field.wire_type() != 2 {
                    return Err(io::Error::other("ambiguous owned-storage field").into());
                }
                let mut owned_storage = shape_field.canonical_payload()?.to_vec();
                litchi_iwa_common::wire::append_varint_field(&mut owned_storage, 99, 9_999)?;
                litchi_iwa_common::wire::append_length_delimited_field(
                    &mut rewritten_shape,
                    4,
                    &owned_storage,
                )?;
            }
            if !selected_storage {
                return Err(io::Error::other("missing owned-storage field").into());
            }
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut rewritten_placeholder,
                1,
                &rewritten_shape,
            )?;
        }
        if !selected_super {
            return Err(io::Error::other("missing placeholder super field").into());
        }
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data: rewritten_placeholder,
            },
        )?;
        Ok(())
    })
}

fn with_unrelated_placeholder(
    package: &[u8],
    owned_storage: Option<u64>,
    deprecated_storage: Option<u64>,
) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let mut payload =
            placeholder_storage_topology(owned_storage, deprecated_storage).encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(&mut payload, 98, 50_198)?;
        let mut placeholder = ArchiveObject::new(
            UNRELATED_PLACEHOLDER,
            vec![
                RawMessage {
                    type_: 777,
                    data: UNRELATED_PLACEHOLDER_MARKER.to_vec(),
                },
                RawMessage {
                    type_: PLACEHOLDER_MESSAGE_TYPE,
                    data: payload,
                },
                RawMessage {
                    type_: 778,
                    data: b"after-unrelated-placeholder".to_vec(),
                },
            ],
        )?;
        let mut references = Vec::with_capacity(2);
        if let Some(identifier) = owned_storage {
            references.push(identifier);
        }
        if let Some(identifier) = deprecated_storage
            && !references.contains(&identifier)
        {
            references.push(identifier);
        }
        placeholder.archive_info.message_infos[1].object_references = references;
        archive.objects.push(placeholder);
        let slide = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        let index = slide
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing slide message"))?;
        let mut value = kn::SlideArchive::decode(slide.messages[index].data.as_slice())?;
        value.slide_number_placeholder = Some(reference(UNRELATED_PLACEHOLDER));
        value.owned_drawables.push(reference(UNRELATED_PLACEHOLDER));
        value
            .drawables_z_order
            .push(reference(UNRELATED_PLACEHOLDER));
        slide.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: value.encode_to_vec(),
            },
        )?;
        slide.archive_info.message_infos[index]
            .object_references
            .push(UNRELATED_PLACEHOLDER);
        Ok(())
    })
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote message"))?;
    Ok(message.data.clone())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Keynote object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let end = start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    assert_eq!(end, data_offset);
    Ok(stream[start..end].to_vec())
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn assert_document_changed_previews_removed_and_other_entries_preserved(
    before: &[u8],
    after: &[u8],
) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog.iter().collect::<Vec<_>>();
    let after_entries = after_catalog.iter().collect::<Vec<_>>();
    assert_eq!(
        before_entries.len(),
        after_entries.len() + PREVIEW_MEMBERS.len()
    );
    let mut changed_document = 0usize;
    let mut removed_previews = 0usize;
    for before_entry in before_entries {
        let matching = after_entries
            .iter()
            .filter(|entry| entry.name() == before_entry.name())
            .collect::<Vec<_>>();
        if PREVIEW_MEMBERS.contains(&before_entry.name()) {
            assert!(matching.is_empty());
            removed_previews += 1;
            continue;
        }
        assert_eq!(matching.len(), 1);
        let after_entry = matching[0];
        assert_eq!(before_entry.raw_name(), after_entry.raw_name());
        if before_entry.name() == DOCUMENT_MEMBER {
            assert_ne!(before_entry.data(), after_entry.data());
            changed_document += 1;
        } else {
            assert_eq!(before_entry.data(), after_entry.data());
            assert_eq!(before_entry.metadata(), after_entry.metadata());
            assert_eq!(
                before_entry.raw_record().local_record(),
                after_entry.raw_record().local_record()
            );
        }
    }
    assert_eq!(changed_document, 1);
    assert_eq!(removed_previews, PREVIEW_MEMBERS.len());
    Ok(())
}

fn assert_root_previews_absent(package: &[u8]) -> TestResult<()> {
    let catalog = Catalog::from_bytes(package)?;
    for member in PREVIEW_MEMBERS {
        assert!(catalog.iter().all(|entry| entry.name() != member));
    }
    Ok(())
}

fn assert_root_previews_restored(before: &[u8], restored: &[u8]) -> TestResult<()> {
    let before = Catalog::from_bytes(before)?;
    let restored = Catalog::from_bytes(restored)?;
    for member in PREVIEW_MEMBERS {
        let source = before
            .iter()
            .find(|entry| entry.name() == member)
            .ok_or_else(|| io::Error::other("source preview is missing"))?;
        let candidate = restored
            .iter()
            .find(|entry| entry.name() == member)
            .ok_or_else(|| io::Error::other("restored preview is missing"))?;
        assert_eq!(candidate.data(), source.data());
        assert_eq!(candidate.metadata(), source.metadata());
        assert_eq!(
            candidate.raw_record().local_record(),
            source.raw_record().local_record()
        );
    }
    Ok(())
}

fn assert_first_slide_node_cache_invalidated(package: &[u8]) -> TestResult<()> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.messages.iter().any(|message| message.type_ == 4))
        .ok_or_else(|| io::Error::other("missing first slide node"))?;
    let index = object
        .messages
        .iter()
        .position(|message| message.type_ == 4)
        .ok_or_else(|| io::Error::other("missing slide-node message"))?;
    let node = kn::SlideNodeArchive::decode(object.messages[index].data.as_slice())?;
    assert!(node.thumbnails.is_empty());
    assert!(node.thumbnail_sizes.is_empty());
    assert!(
        node.digests_for_datas_needing_download_for_thumbnail
            .is_empty()
    );
    assert_eq!(node.thumbnails_are_dirty, Some(true));
    assert!(
        object.archive_info.message_infos[index]
            .data_references
            .is_empty()
    );
    assert!(
        object.archive_info.message_infos[index]
            .field_infos
            .iter()
            .all(|field| field.data_references.is_empty())
    );
    Ok(())
}

fn assert_first_slide_node_restored(before: &[u8], restored: &[u8]) -> TestResult<()> {
    let before_stream = document_stream(before)?;
    let before_archive = Archive::parse(&before_stream)?;
    let before = before_archive
        .objects
        .iter()
        .find(|object| object.messages.iter().any(|message| message.type_ == 4))
        .ok_or_else(|| io::Error::other("source slide node is missing"))?;
    let restored_stream = document_stream(restored)?;
    let restored_archive = Archive::parse(&restored_stream)?;
    let restored = restored_archive
        .objects
        .iter()
        .find(|object| object.messages.iter().any(|message| message.type_ == 4))
        .ok_or_else(|| io::Error::other("restored slide node is missing"))?;
    assert_eq!(restored.messages, before.messages);
    assert_eq!(restored.archive_info, before.archive_info);
    Ok(())
}

fn native_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

#[test]
fn title_body_reads_are_role_distinct_and_preserve_absent_vs_empty() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.slide_text("Agenda", SlideTextRole::Title)?,
        Some(TITLE.to_owned())
    );
    assert_eq!(
        package.slide_title(SlideSelector::index(0))?,
        Some(TITLE.to_owned())
    );
    assert_eq!(
        package.slide_text("Agenda", SlideTextRole::Body)?,
        Some(BODY.to_owned())
    );
    assert_eq!(
        package.slide_body(SlideSelector::index(0))?,
        Some(BODY.to_owned())
    );
    assert_eq!(package.slide_title("Existing Empty")?, Some(String::new()));
    assert_eq!(package.slide_body("Existing Empty")?, None);
    assert_ne!(package.show()?.slides()[0].name(), Some(TITLE));

    assert!(matches!(
        package.edit_slide_body("Existing Empty"),
        Err(SlideTextError::TextStorageNotFound {
            role: SlideTextRole::Body
        })
    ));
    Ok(())
}

#[test]
fn native_fixture_exposes_existing_title_and_body_without_raw_identity() -> TestResult<()> {
    let package = Package::open(native_fixture_path())?;
    let title = package
        .slide_title(SlideSelector::index(0))?
        .ok_or_else(|| io::Error::other("native fixture has no title storage"))?;
    let body = package
        .slide_body(SlideSelector::index(0))?
        .ok_or_else(|| io::Error::other("native fixture has no body storage"))?;
    assert!(title.contains("Litchi native Keynote fixture"));
    assert_eq!(body, "Buffa lazy-view migration verification");
    assert!(!body.contains("2026-08-07"));
    assert!(package.text()?.contains("2026-08-07"));
    Ok(())
}

#[test]
fn native_fixture_changed_title_commits_and_inverse_restores_exact_bytes() -> TestResult<()> {
    let bytes = std::fs::read(native_fixture_path())?;
    let package = Package::from_bytes(&bytes)?;
    let before = package
        .slide_title(SlideSelector::index(0))?
        .ok_or_else(|| io::Error::other("native fixture has no title storage"))?;
    let body = package.slide_body(SlideSelector::index(0))?;

    let replacement = "Litchi native Keynote fixture — changed 🚀";
    let mut edit = package.edit_slide_title(SlideSelector::index(0))?;
    edit.set(replacement)?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(
        commit.package().slide_title(SlideSelector::index(0))?,
        Some(replacement.to_owned())
    );
    assert_eq!(commit.package().slide_body(SlideSelector::index(0))?, body);
    assert_ne!(commit.package().source_bytes(), bytes);
    assert_root_previews_absent(commit.package().source_bytes())?;
    assert_first_slide_node_cache_invalidated(commit.package().source_bytes())?;
    assert!(commit.package().text()?.contains("2026-08-07"));

    let restored = commit
        .package()
        .apply_slide_text(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_root_previews_restored(&bytes, restored.package().source_bytes())?;
    assert_first_slide_node_restored(&bytes, restored.package().source_bytes())?;
    assert_eq!(
        restored.package().slide_title(SlideSelector::index(0))?,
        Some(before)
    );
    Ok(())
}

#[test]
fn deprecated_storage_must_duplicate_the_mandatory_owned_storage() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let matching = with_title_storage_topology(&bytes, Some(TITLE_STORAGE), Some(TITLE_STORAGE))?;
    let matching_package = Package::from_bytes(&matching)?;
    assert_eq!(
        matching_package.slide_title("Agenda")?,
        Some(TITLE.to_owned())
    );
    let mut edit = matching_package.edit_slide_title("Agenda")?;
    edit.set("Matching legacy and mandatory title storage")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().slide_title("Agenda")?,
        Some("Matching legacy and mandatory title storage".to_owned())
    );
    assert_eq!(
        commit
            .package()
            .apply_slide_text(&commit.patch().inverse())?
            .package()
            .source_bytes(),
        matching
    );

    let mismatch = with_title_storage_topology(&bytes, Some(TITLE_STORAGE), Some(BODY_STORAGE))?;
    let mismatch_package = Package::from_bytes(&mismatch)?;
    assert!(matches!(
        mismatch_package.edit_slide_title("Agenda"),
        Err(SlideTextError::DependentContent | SlideTextError::InvalidSource)
    ));
    assert_eq!(mismatch_package.source_bytes(), mismatch);

    let deprecated_only = with_title_storage_topology(&bytes, None, Some(TITLE_STORAGE))?;
    let deprecated_only_package = Package::from_bytes(&deprecated_only)?;
    assert!(matches!(
        deprecated_only_package.edit_slide_title("Agenda"),
        Err(SlideTextError::InvalidSource)
    ));
    assert_eq!(deprecated_only_package.source_bytes(), deprecated_only);
    Ok(())
}

#[test]
fn unrelated_zero_placeholder_is_invisible_and_preserved_by_title_and_body_commits()
-> TestResult<()> {
    let bytes = with_unrelated_placeholder(&synthetic_package(false)?, Some(0), Some(0))?;
    let source_payload = message_payload(&bytes, UNRELATED_PLACEHOLDER, PLACEHOLDER_MESSAGE_TYPE)?;
    let source_before = message_payload(&bytes, UNRELATED_PLACEHOLDER, 777)?;
    let source_after = message_payload(&bytes, UNRELATED_PLACEHOLDER, 778)?;
    let source_header = object_header(&bytes, UNRELATED_PLACEHOLDER)?;
    let marker = std::str::from_utf8(UNRELATED_PLACEHOLDER_MARKER)?;

    for (role, other_role, replacement) in [
        (
            SlideTextRole::Title,
            SlideTextRole::Body,
            "Selected title changed beside zero sentinel",
        ),
        (
            SlideTextRole::Body,
            SlideTextRole::Title,
            "Selected body changed beside zero sentinel",
        ),
    ] {
        let package = Package::from_bytes(&bytes)?;
        assert!(!package.text()?.contains(marker));
        let other_before = package.slide_text("Agenda", other_role)?;
        let mut edit = package.edit_slide_text("Agenda", role)?;
        edit.set(replacement)?;
        let commit = edit.commit()?;
        assert_eq!(
            commit.package().slide_text("Agenda", role)?,
            Some(replacement.to_owned())
        );
        assert_eq!(
            commit.package().slide_text("Agenda", other_role)?,
            other_before
        );
        assert!(!commit.package().text()?.contains(marker));
        assert_eq!(
            message_payload(
                commit.package().source_bytes(),
                UNRELATED_PLACEHOLDER,
                PLACEHOLDER_MESSAGE_TYPE,
            )?,
            source_payload
        );
        assert_eq!(
            message_payload(commit.package().source_bytes(), UNRELATED_PLACEHOLDER, 777)?,
            source_before
        );
        assert_eq!(
            message_payload(commit.package().source_bytes(), UNRELATED_PLACEHOLDER, 778)?,
            source_after
        );
        assert_eq!(
            object_header(commit.package().source_bytes(), UNRELATED_PLACEHOLDER)?,
            source_header
        );
    }
    Ok(())
}

#[test]
fn unrelated_placeholder_cannot_alias_selected_storage_by_modern_or_legacy_edge() -> TestResult<()>
{
    let bytes = synthetic_package(false)?;
    for adversarial in [
        with_unrelated_placeholder(&bytes, Some(TITLE_STORAGE), None)?,
        with_unrelated_placeholder(&bytes, Some(0), Some(TITLE_STORAGE))?,
    ] {
        let package = Package::from_bytes(&adversarial)?;
        let mut edit = package.edit_slide_title("Agenda")?;
        edit.set("must not publish across an unrelated owner")?;
        assert!(matches!(
            edit.commit(),
            Err(SlideTextError::DependentContent)
        ));
        assert_eq!(package.source_bytes(), adversarial);
    }
    Ok(())
}

#[test]
fn selector_first_utf16_operations_cover_both_roles() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let untouched_body = package.slide_body("Agenda")?;
    let untouched_sibling = message_payload(&bytes, SIBLING_STORAGE, STORAGE_MESSAGE_TYPE)?;

    let span = TextSpan::from_utf16_indexes(7, 9)?;
    let mut edit = package.edit_slide_title("Agenda")?;
    assert_eq!(edit.position(), Position::new(0));
    assert_eq!(edit.role(), SlideTextRole::Title);
    assert_eq!(edit.text(), TITLE);
    edit.replace(span, "東京😀")?;
    assert_eq!(edit.span(), Some(span));
    let commit = edit.commit()?;
    assert_eq!(commit.patch().position(), Position::new(0));
    assert_eq!(commit.patch().role(), SlideTextRole::Title);
    assert_eq!(commit.patch().span(), span);
    assert_eq!(commit.patch().before(), TITLE);
    assert_eq!(commit.patch().after(), "Launch 東京😀 title");
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit.package().slide_title("Agenda")?,
        Some("Launch 東京😀 title".to_owned())
    );
    assert_eq!(commit.package().slide_body("Agenda")?, untouched_body);
    assert_eq!(
        message_payload(
            commit.package().source_bytes(),
            SIBLING_STORAGE,
            STORAGE_MESSAGE_TYPE
        )?,
        untouched_sibling
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let mut insert = package.edit_slide_body("Agenda")?;
    insert.insert(TextPosition::ZERO, "Intro: ")?;
    assert_eq!(
        insert.commit()?.package().slide_body("Agenda")?,
        Some("Intro: Body 東京😀".to_owned())
    );

    let mut delete = package.edit_slide_body("Agenda")?;
    delete.delete(TextSpan::from_utf16_indexes(5, 7)?)?;
    assert_eq!(
        delete.commit()?.package().slide_body("Agenda")?,
        Some("Body 😀".to_owned())
    );

    let mut set = package.edit_slide_title("Agenda")?;
    set.set("Entirely new title")?;
    assert_eq!(
        set.commit()?.package().slide_title("Agenda")?,
        Some("Entirely new title".to_owned())
    );

    let mut clear = package.edit_slide_body("Agenda")?;
    clear.clear()?;
    let cleared = clear.commit()?;
    assert_eq!(cleared.package().slide_body("Agenda")?, Some(String::new()));
    let mut refill = cleared.package().edit_slide_body("Agenda")?;
    assert_eq!(refill.text(), "");
    refill.set("Restored body")?;
    assert_eq!(
        refill.commit()?.package().slide_body("Agenda")?,
        Some("Restored body".to_owned())
    );
    Ok(())
}

#[test]
fn utf16_boundaries_selector_and_staging_errors_leave_source_unchanged() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();

    assert!(matches!(
        package.edit_slide_title("Missing"),
        Err(SlideTextError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.edit_slide_text(SlideSelector::index(2), SlideTextRole::Title),
        Err(SlideTextError::SlidePositionNotFound { position }) if position == Position::new(2)
    ));

    let mut split_start = package.edit_slide_title("Agenda")?;
    assert!(matches!(
        split_start.replace(TextSpan::from_utf16_indexes(8, 9)?, "x"),
        Err(SlideTextError::SurrogateBoundary { position })
            if position == TextPosition::from_utf16_code_units(8)
    ));
    let mut split_end = package.edit_slide_title("Agenda")?;
    assert!(matches!(
        split_end.delete(TextSpan::from_utf16_indexes(7, 8)?),
        Err(SlideTextError::SurrogateBoundary { position })
            if position == TextPosition::from_utf16_code_units(8)
    ));
    let mut out_of_bounds = package.edit_slide_body("Agenda")?;
    assert!(matches!(
        out_of_bounds.delete(TextSpan::from_utf16_indexes(0, 100)?),
        Err(SlideTextError::SpanOutOfBounds { length, .. })
            if length == TextPosition::from_utf16_code_units(9)
    ));
    let mut marker = package.edit_slide_title("Agenda")?;
    assert!(matches!(
        marker.insert(TextPosition::ZERO, "bad\u{fffc}marker"),
        Err(SlideTextError::ObjectMarkerReplacement)
    ));
    let mut one_operation = package.edit_slide_body("Agenda")?;
    one_operation.insert(TextPosition::ZERO, "first")?;
    assert!(matches!(
        one_operation.insert(TextPosition::ZERO, "second"),
        Err(SlideTextError::OperationAlreadyStaged)
    ));
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn semantic_noops_share_the_exact_source_allocation_for_each_role() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();

    for (role, text) in [(SlideTextRole::Title, TITLE), (SlideTextRole::Body, BODY)] {
        let mut edit = package.edit_slide_text("Agenda", role)?;
        edit.set(text)?;
        let commit = edit.commit()?;
        assert_eq!(commit.patch().role(), role);
        assert!(commit.patch().is_noop());
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.package().source_bytes(), bytes);
        assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);
        let applied = package.apply_slide_text(commit.patch())?;
        assert!(applied.patch().is_noop());
        assert_eq!(applied.package().source_bytes().as_ptr(), source_pointer);
    }

    let unstaged = package.edit_slide_title("Agenda")?.commit()?;
    assert!(unstaged.patch().is_noop());
    assert_eq!(unstaged.package().source_bytes().as_ptr(), source_pointer);
    Ok(())
}

#[test]
fn empty_unicode_and_large_values_round_trip_and_reverse_exactly() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let large = "🙂e\u{301} שלום 東京\n".repeat(2_048);

    for replacement in ["", "\0naïve — مرحبا — 👩🏽‍🚀", large.as_str()] {
        let package = Package::from_bytes(&bytes)?;
        let mut edit = package.edit_slide_title("Agenda")?;
        edit.set(replacement)?;
        let commit = edit.commit()?;

        assert!(commit.diagnostics().changed());
        assert_eq!(commit.patch().before(), TITLE);
        assert_eq!(commit.patch().after(), replacement);
        assert_eq!(
            commit.package().slide_title("Agenda")?,
            Some(replacement.to_owned())
        );
        assert_eq!(
            package
                .apply_slide_text(commit.patch())?
                .package()
                .source_bytes(),
            commit.package().source_bytes()
        );
        assert_eq!(
            commit
                .package()
                .apply_slide_text(&commit.patch().inverse())?
                .package()
                .source_bytes(),
            bytes
        );
    }
    Ok(())
}

#[test]
fn stale_and_replayed_patches_fail_with_patch_conflict() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;

    let mut first = package.edit_slide_title("Agenda")?;
    first.set("first concurrent title")?;
    let first = first.commit()?;

    let mut second = package.edit_slide_title("Agenda")?;
    second.set("second concurrent title")?;
    let second = second.commit()?;

    assert!(matches!(
        first.package().apply_slide_text(second.patch()),
        Err(SlideTextError::PatchConflict)
    ));
    assert!(matches!(
        first.package().apply_slide_text(first.patch()),
        Err(SlideTextError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_slide_text(&first.patch().inverse()),
        Err(SlideTextError::PatchConflict)
    ));
    assert_eq!(
        first
            .package()
            .apply_slide_text(&first.patch().inverse())?
            .package()
            .source_bytes(),
        bytes
    );
    Ok(())
}

#[test]
fn changed_text_preserves_unknowns_scope_siblings_and_exact_inverse() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let source_owner = message_payload(&bytes, TITLE_OWNER, PLACEHOLDER_MESSAGE_TYPE)?;
    let source_storage = message_payload(&bytes, TITLE_STORAGE, STORAGE_MESSAGE_TYPE)?;
    let source_before_storage = message_payload(&bytes, TITLE_STORAGE, 779)?;
    let source_after_storage = message_payload(&bytes, TITLE_STORAGE, 780)?;
    let source_header = object_header(&bytes, TITLE_STORAGE)?;
    let source_body = message_payload(&bytes, BODY_STORAGE, STORAGE_MESSAGE_TYPE)?;
    let source_sibling = message_payload(&bytes, SIBLING_STORAGE, STORAGE_MESSAGE_TYPE)?;

    let mut edit = package.edit_slide_title("Agenda")?;
    edit.replace(TextSpan::from_utf16_indexes(7, 9)?, "東京😀")?;
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("Launch"));
    assert!(!edit_debug.contains("Agenda"));
    assert!(!edit_debug.contains("Index/"));
    let commit = edit.commit()?;
    let target = commit.package().source_bytes();
    assert_document_changed_previews_removed_and_other_entries_preserved(&bytes, target)?;
    assert_first_slide_node_cache_invalidated(target)?;
    assert_eq!(
        message_payload(target, TITLE_OWNER, PLACEHOLDER_MESSAGE_TYPE)?,
        source_owner
    );
    assert_eq!(
        message_payload(target, TITLE_STORAGE, 779)?,
        source_before_storage
    );
    assert_eq!(
        message_payload(target, TITLE_STORAGE, 780)?,
        source_after_storage
    );
    assert_eq!(
        message_payload(target, BODY_STORAGE, STORAGE_MESSAGE_TYPE)?,
        source_body
    );
    assert_eq!(
        message_payload(target, SIBLING_STORAGE, STORAGE_MESSAGE_TYPE)?,
        source_sibling
    );
    let target_storage = message_payload(target, TITLE_STORAGE, STORAGE_MESSAGE_TYPE)?;
    assert_eq!(
        raw_fields(&target_storage, 98)?,
        raw_fields(&source_storage, 98)?
    );
    let target_header = object_header(target, TITLE_STORAGE)?;
    assert_eq!(
        raw_fields(&target_header, 99)?,
        raw_fields(&source_header, 99)?
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let applied = package.apply_slide_text(commit.patch())?;
    assert_eq!(applied.package().source_bytes(), target);
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.role(), SlideTextRole::Title);
    assert_eq!(inverse.before(), commit.patch().after());
    assert_eq!(inverse.after(), commit.patch().before());
    assert_eq!(
        inverse.source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(
        inverse.target_fingerprint(),
        commit.patch().source_fingerprint()
    );
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = commit.package().apply_slide_text(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_root_previews_restored(&bytes, restored.package().source_bytes())?;
    assert_first_slide_node_restored(&bytes, restored.package().source_bytes())?;
    assert_eq!(
        restored.package().slide_title("Agenda")?,
        Some(TITLE.to_owned())
    );

    let catalog = Catalog::from_bytes(&bytes)?;
    let equivalent_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"different unrelated ZIP sentinel",
        )],
        Limits::default(),
    )?;
    let equivalent = Package::from_bytes(&equivalent_bytes)?;
    assert_eq!(
        equivalent.slide_title("Agenda")?,
        package.slide_title("Agenda")?
    );
    assert!(matches!(
        equivalent.apply_slide_text(commit.patch()),
        Err(SlideTextError::PatchConflict)
    ));

    let patch_debug = format!("{:?}", commit.patch());
    assert!(!patch_debug.contains("Launch"));
    assert!(!patch_debug.contains("Agenda"));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains(std::str::from_utf8(PRIVATE_MARKER)?));
    assert_send_sync(&package);
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(commit.diagnostics());
    assert_type_send_sync::<SlideTextCommit>();
    assert_type_send_sync::<SlideTextPatch>();
    assert_type_send_sync::<SlideTextDiagnostics>();
    assert_type_send_sync::<SlideTextEdit<'static>>();
    assert_type_send_sync::<SlideTextError>();
    Ok(())
}

#[test]
fn canonical_unknowns_nested_in_owned_storage_references_are_preserved() -> TestResult<()> {
    let bytes = with_owned_storage_reference_unknown(&synthetic_package(false)?, TITLE_OWNER)?;
    let source_owner = message_payload(&bytes, TITLE_OWNER, PLACEHOLDER_MESSAGE_TYPE)?;
    let package = Package::from_bytes(&bytes)?;
    let mut edit = package.edit_slide_title("Agenda")?;
    edit.set("title beside a future reference field")?;
    let commit = edit.commit()?;

    assert_eq!(
        message_payload(
            commit.package().source_bytes(),
            TITLE_OWNER,
            PLACEHOLDER_MESSAGE_TYPE
        )?,
        source_owner
    );
    assert_eq!(
        commit
            .package()
            .apply_slide_text(&commit.patch().inverse())?
            .package()
            .source_bytes(),
        bytes
    );
    Ok(())
}

#[test]
fn duplicate_slide_names_are_typed_as_ambiguous_without_affecting_positions() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let duplicate_name = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(SECOND_SLIDE)
            .ok_or_else(|| io::Error::other("missing second slide"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing second slide message"))?;
        let mut value = kn::SlideArchive::decode(object.messages[index].data.as_slice())?;
        value.name = Some("Agenda".to_owned());
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: value.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&duplicate_name)?;

    assert!(matches!(
        package.slide_title("Agenda"),
        Err(SlideTextError::AmbiguousSelector)
    ));
    assert!(matches!(
        package.edit_slide_body("Agenda"),
        Err(SlideTextError::AmbiguousSelector)
    ));
    assert_eq!(
        package.slide_title(SlideSelector::index(0))?,
        Some(TITLE.to_owned())
    );
    assert_eq!(
        package.slide_title(SlideSelector::index(1))?,
        Some(String::new())
    );
    assert_eq!(package.source_bytes(), duplicate_name);
    Ok(())
}

#[test]
fn explicit_placeholder_kinds_are_role_checked_for_both_roles() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    for (owner, role, allowed, rejected) in [
        (
            TITLE_OWNER,
            SlideTextRole::Title,
            [None, Some(0), Some(2)],
            [1, 3, 4, 99],
        ),
        (
            BODY_OWNER,
            SlideTextRole::Body,
            [None, Some(0), Some(3)],
            [1, 2, 4, 99],
        ),
    ] {
        for kind in allowed {
            let compatible = with_placeholder_kind(&bytes, owner, kind)?;
            let package = Package::from_bytes(&compatible)?;
            let mut edit = package.edit_slide_text("Agenda", role)?;
            edit.set(match role {
                SlideTextRole::Title => "compatible title kind",
                SlideTextRole::Body => "compatible body kind",
                _ => "compatible placeholder kind",
            })?;
            let commit = edit.commit()?;
            assert_eq!(
                commit
                    .package()
                    .apply_slide_text(&commit.patch().inverse())?
                    .package()
                    .source_bytes(),
                compatible
            );
        }
        for kind in rejected {
            let contradictory = with_placeholder_kind(&bytes, owner, Some(kind))?;
            let package = Package::from_bytes(&contradictory)?;
            assert!(matches!(
                package.edit_slide_text("Agenda", role),
                Err(SlideTextError::DependentContent)
            ));
            assert_eq!(package.source_bytes(), contradictory);
        }
    }
    Ok(())
}

#[test]
fn duplicate_body_payload_and_field_path_ownership_fail_closed() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let duplicate_body_payload = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing slide message"))?;
        let mut data = object.messages[index].data.clone();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut data,
            6,
            &reference(BODY_OWNER).encode_to_vec(),
        )?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&duplicate_body_payload)?;
    assert!(matches!(
        package.edit_slide_body("Agenda"),
        Err(SlideTextError::InvalidSource)
    ));
    assert_eq!(package.source_bytes(), duplicate_body_payload);

    let duplicate_field_path = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        let mut field = FieldInfo::new(vec![6]);
        field.object_references.push(BODY_OWNER);
        object.archive_info.message_infos[0]
            .field_infos
            .extend([field.clone(), field]);
        Ok(())
    })?;
    let package = Package::from_bytes(&duplicate_field_path)?;
    let mut edit = package.edit_slide_body("Agenda")?;
    edit.set("duplicate metadata must not publish")?;
    assert!(matches!(
        edit.commit(),
        Err(SlideTextError::DependentContent)
    ));
    assert_eq!(package.source_bytes(), duplicate_field_path);
    Ok(())
}

#[test]
fn selected_storage_merge_and_diff_metadata_refuse_changed_commits() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    for mutation in 0..6 {
        let adversarial = rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_STORAGE)
                .ok_or_else(|| io::Error::other("missing title storage"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?;
            let info = &mut object.archive_info.message_infos[index];
            match mutation {
                0 => object.archive_info.should_merge = Some(true),
                1 => info.base_message_index = Some(0),
                2 => info.diff_merge_version.push(1),
                3 => info.diff_field_path = Some(vec![1].into()),
                4 => info.fields_to_remove.push(vec![1].into()),
                5 => info.diff_read_version.push(1),
                _ => return Err(io::Error::other("unknown metadata mutation").into()),
            }
            Ok(())
        })?;
        let package = Package::from_bytes(&adversarial)?;
        let mut edit = package.edit_slide_title("Agenda")?;
        edit.set("merge metadata must remain opaque")?;
        assert!(matches!(
            edit.commit(),
            Err(SlideTextError::DependentContent)
        ));
        assert_eq!(package.source_bytes(), adversarial);
    }
    Ok(())
}

#[test]
fn malformed_and_aliased_role_graphs_fail_closed() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let malformed = [
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(FIRST_SLIDE)
                .ok_or_else(|| io::Error::other("missing first slide"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing slide message"))?;
            let mut data = object.messages[index].data.clone();
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut data,
                5,
                &reference(TITLE_OWNER).encode_to_vec(),
            )?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: SLIDE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(TITLE_OWNER));
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_OWNER)
                .ok_or_else(|| io::Error::other("missing title owner"))?;
            object.push_message(object.messages[0].clone())?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_OWNER)
                .ok_or_else(|| io::Error::other("missing title owner"))?;
            object.replace_message_preserving_header(
                0,
                RawMessage {
                    type_: PLACEHOLDER_MESSAGE_TYPE,
                    data: Vec::new(),
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(TITLE_STORAGE));
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_STORAGE)
                .ok_or_else(|| io::Error::other("missing title storage"))?;
            let duplicate = object
                .messages
                .iter()
                .find(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?
                .clone();
            object.push_message(duplicate)?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_STORAGE)
                .ok_or_else(|| io::Error::other("missing title storage"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: STORAGE_MESSAGE_TYPE,
                    data: vec![0x1a, 0x80],
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_STORAGE)
                .ok_or_else(|| io::Error::other("missing title storage"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: STORAGE_MESSAGE_TYPE,
                    data: vec![0x1a, 0x01, 0xff],
                },
            )?;
            Ok(())
        })?,
        rewrite_document(&bytes, |archive| {
            let object = archive
                .object_mut(TITLE_STORAGE)
                .ok_or_else(|| io::Error::other("missing title storage"))?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing storage message"))?;
            object.replace_message_preserving_header(
                index,
                RawMessage {
                    type_: STORAGE_MESSAGE_TYPE,
                    data: vec![0x18, 0x01],
                },
            )?;
            Ok(())
        })?,
    ];

    for malformed_bytes in malformed {
        let package = Package::from_bytes(&malformed_bytes)?;
        let result = package.edit_slide_title(SlideSelector::index(0));
        assert!(
            matches!(result, Err(SlideTextError::InvalidSource)),
            "unexpected malformed slide-text result: {result:?}"
        );
        assert_eq!(package.source_bytes(), malformed_bytes);
    }

    let alias = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(BODY_OWNER)
            .ok_or_else(|| io::Error::other("missing body owner"))?;
        object.replace_message_preserving_header(
            0,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data: placeholder(TITLE_STORAGE).encode_to_vec(),
            },
        )?;
        object.archive_info.message_infos[0].object_references = vec![TITLE_STORAGE];
        Ok(())
    })?;
    let package = Package::from_bytes(&alias)?;
    let mut edit = package.edit_slide_title("Agenda")?;
    edit.set("must not rewrite aliased storage")?;
    assert!(matches!(
        edit.commit(),
        Err(SlideTextError::DependentContent)
    ));
    assert_eq!(package.source_bytes(), alias);
    Ok(())
}

#[test]
fn exact_noop_commit_and_apply_do_not_require_mutation_ownership() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let malformed_metadata = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        object.archive_info.message_infos[0]
            .object_references
            .retain(|identifier| *identifier != TITLE_OWNER);
        Ok(())
    })?;
    let aliased_storage = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(BODY_OWNER)
            .ok_or_else(|| io::Error::other("missing body owner"))?;
        object.replace_message_preserving_header(
            0,
            RawMessage {
                type_: PLACEHOLDER_MESSAGE_TYPE,
                data: placeholder(TITLE_STORAGE).encode_to_vec(),
            },
        )?;
        object.archive_info.message_infos[0].object_references = vec![TITLE_STORAGE];
        Ok(())
    })?;

    for adversarial in [malformed_metadata, aliased_storage] {
        let package = Package::from_bytes(&adversarial)?;
        let source_pointer = package.source_bytes().as_ptr();
        let mut edit = package.edit_slide_title("Agenda")?;
        edit.set(TITLE)?;
        let commit = edit.commit()?;
        assert!(commit.patch().is_noop());
        assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);
        let replay = package.apply_slide_text(commit.patch())?;
        assert!(replay.patch().is_noop());
        assert_eq!(replay.package().source_bytes().as_ptr(), source_pointer);
    }
    Ok(())
}

#[test]
fn unrelated_noncanonical_object_prefix_is_refused_before_rewrite() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let noncanonical = with_overlong_object_length_prefix(&bytes, 500)?;
    let package = Package::from_bytes(&noncanonical)?;
    let mut edit = package.edit_slide_title("Agenda")?;
    edit.set("must not canonicalize an unrelated object frame")?;
    assert!(matches!(edit.commit(), Err(SlideTextError::InvalidSource)));
    assert_eq!(package.source_bytes(), noncanonical);
    Ok(())
}

#[test]
fn ownership_metadata_and_dependent_text_are_proven_before_publish() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let missing_slide_ownership = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(FIRST_SLIDE)
            .ok_or_else(|| io::Error::other("missing first slide"))?;
        object.archive_info.message_infos[0]
            .object_references
            .retain(|identifier| *identifier != TITLE_OWNER);
        Ok(())
    })?;
    let missing_owner_ownership = rewrite_document(&bytes, |archive| {
        let object = archive
            .object_mut(TITLE_OWNER)
            .ok_or_else(|| io::Error::other("missing title owner"))?;
        object.archive_info.message_infos[0]
            .object_references
            .clear();
        Ok(())
    })?;
    for malformed_bytes in [missing_slide_ownership, missing_owner_ownership] {
        let package = Package::from_bytes(&malformed_bytes)?;
        let mut edit = package.edit_slide_title("Agenda")?;
        edit.set("ownership must be proven")?;
        assert!(matches!(edit.commit(), Err(SlideTextError::InvalidSource)));
        assert_eq!(package.source_bytes(), malformed_bytes);
    }

    let dependent = synthetic_package(true)?;
    let package = Package::from_bytes(&dependent)?;
    let mut intersects = package.edit_slide_title("Agenda")?;
    intersects.delete(TextSpan::from_utf16_indexes(7, 9)?)?;
    assert!(matches!(
        intersects.commit(),
        Err(SlideTextError::DependentContent)
    ));
    let mut unrelated = package.edit_slide_title("Agenda")?;
    unrelated.replace(TextSpan::from_utf16_indexes(0, 1)?, "X")?;
    assert_eq!(
        unrelated.commit()?.package().slide_title("Agenda")?,
        Some("Xaunch 🚀 title".to_owned())
    );
    Ok(())
}

#[test]
fn errors_and_transaction_debug_are_content_free() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    let mut edit = package.edit_slide_body("Agenda")?;
    edit.insert(TextPosition::ZERO, "authored secret")?;
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("authored secret"));
    assert!(!edit_debug.contains(BODY));
    assert!(!edit_debug.contains("Agenda"));

    let errors = [
        SlideTextError::SlideNameNotFound,
        SlideTextError::TextStorageNotFound {
            role: SlideTextRole::Title,
        },
        SlideTextError::PatchConflict,
        SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::OutputBytes,
            observed: 11,
            maximum: 10,
        },
    ];
    for error in errors {
        let error_debug = format!("{error:?}");
        let display = error.to_string();
        for secret in [TITLE, BODY, SIBLING, "Agenda", "Index/", "authored secret"] {
            assert!(!error_debug.contains(secret));
            assert!(!display.contains(secret));
        }
    }
    Ok(())
}

#[test]
fn selected_text_limit_is_inclusive_and_one_over_is_typed() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let semantic = SemanticLimits::default();
    let exact = SemanticLimits::new(
        semantic.max_objects(),
        semantic.max_slides(),
        semantic.max_references(),
        semantic.max_text_storages(),
        semantic.max_text_fragments(),
        TITLE.len(),
    )?;
    let exact_package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), exact))?;
    assert_eq!(
        exact_package.slide_title(SlideSelector::index(0))?,
        Some(TITLE.to_owned())
    );

    let one_under = SemanticLimits::new(
        semantic.max_objects(),
        semantic.max_slides(),
        semantic.max_references(),
        semantic.max_text_storages(),
        semantic.max_text_fragments(),
        TITLE.len() - 1,
    )?;
    let limited_package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), one_under))?;
    assert!(matches!(
        limited_package.slide_title(SlideSelector::index(0)),
        Err(SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::TextBytes,
            observed,
            maximum,
        }) if observed == u64::try_from(TITLE.len())?
            && maximum == u64::try_from(TITLE.len() - 1)?
    ));
    Ok(())
}

#[test]
fn candidate_text_limit_is_checked_when_staging_with_exact_amounts() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let semantic = SemanticLimits::default();
    let limits = SemanticLimits::new(
        semantic.max_objects(),
        semantic.max_slides(),
        semantic.max_references(),
        semantic.max_text_storages(),
        semantic.max_text_fragments(),
        TITLE.len(),
    )?;
    let package =
        Package::from_bytes_with_options(&bytes, ReadOptions::new(Limits::default(), limits))?;
    let mut edit = package.edit_slide_title(SlideSelector::index(0))?;
    let error = edit
        .insert(TextPosition::ZERO, "x")
        .err()
        .ok_or_else(|| io::Error::other("over-limit candidate should be rejected"))?;

    assert!(matches!(
        error,
        SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::TextBytes,
            observed,
            maximum,
        } if observed == u64::try_from(TITLE.len() + 1)?
            && maximum == u64::try_from(TITLE.len())?
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn changed_slide_text_respects_the_retained_output_limit() -> TestResult<()> {
    let bytes = synthetic_package(false)?;
    let input_bytes = u64::try_from(bytes.len())?;
    let limits = Limits::new(input_bytes, 16, 1024 * 1024, 1024 * 1024, 1024 * 1024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut edit = package.edit_slide_title("Agenda")?;
    let mut expanded = String::with_capacity(128 * 1_024);
    let mut state = 0x5eed_1234_89ab_cdef_u64;
    for _ in 0..(128 * 1_024) {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        let byte = b'!'.saturating_add(u8::try_from(state % 94)?);
        expanded.push(char::from(byte));
    }
    edit.insert(TextPosition::ZERO, &expanded)?;
    assert!(matches!(
        edit.commit(),
        Err(SlideTextError::LimitExceeded {
            kind: SlideTextLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}
