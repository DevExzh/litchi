use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsk, tsp, tswp};
use litchi_pages::{
    Limits, Package, PageLayoutCommit, PageLayoutDiagnostics, PageLayoutEdit, PageLayoutError,
    PageLayoutLimitKind, PageLayoutPatch,
    page_layout::{Layout, Orientation},
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const DOCUMENT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const SECTION_IDENTIFIER: u64 = 43;
const VIEW_STATE_IDENTIFIER: u64 = 49;
const VIEW_STATE_ROOT_IDENTIFIER: u64 = 50;
const LAYOUT_STATE_IDENTIFIER: u64 = 51;
const DETACHED_VIEW_STATE_ROOT_IDENTIFIER: u64 = 60;
const DETACHED_LAYOUT_STATE_IDENTIFIER: u64 = 61;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const VIEW_STATE_ROOT_MESSAGE_TYPE: u32 = 10_147;
const PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const PRIVATE_MARKER: &[u8] = b"private-pages-page-layout-marker-998244353";

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

fn fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/basic.pages")
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
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
    let (header_length, prefix_length) = decode_varint_from_bytes(&bytes[header_offset..])?;
    let header_start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let header_end = header_start
        .checked_add(usize::try_from(header_length)?)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    assert_eq!(header_end, data_offset);
    let mut header = bytes[header_start..header_end].to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut rewritten = Vec::with_capacity(bytes.len().saturating_add(8));
    rewritten.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&bytes[data_offset..]);
    assert_eq!(Archive::parse(&rewritten)?.to_bytes()?, rewritten);
    Ok(SnappyStream::compress(&rewritten)?)
}

fn full_layout() -> TestResult<Layout> {
    Ok(Layout::new(
        Some(612.0),
        Some(792.0),
        Some(72.0),
        Some(72.0),
        Some(54.0),
        Some(54.0),
        Some(36.0),
        Some(36.0),
        Some(1.0),
        Some(Orientation::Portrait),
        Some(false),
    )?)
}

fn document_payload(layout: Layout, view_state: Option<u64>) -> TestResult<Vec<u8>> {
    let mut payload = tp::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            view_state: view_state.map(reference),
            ..Default::default()
        },
        body_storage: Some(reference(BODY_IDENTIFIER)),
        section: Some(reference(SECTION_IDENTIFIER)),
        page_width: layout.page_width(),
        page_height: layout.page_height(),
        left_margin: layout.left_margin(),
        right_margin: layout.right_margin(),
        top_margin: layout.top_margin(),
        bottom_margin: layout.bottom_margin(),
        header_margin: layout.header_margin(),
        footer_margin: layout.footer_margin(),
        page_scale: layout.page_scale(),
        lays_out_body_vertically: layout.lays_out_body_vertically(),
        orientation: layout.orientation().map(Orientation::as_raw),
        ..Default::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut payload, 99, 123_456)?;
    Ok(payload)
}

fn synthetic_package(
    layout: Layout,
    include_view_state: bool,
    include_previews: bool,
) -> TestResult<Vec<u8>> {
    synthetic_package_with_view_location(layout, include_view_state, include_previews, false)
}

fn synthetic_package_with_view_location(
    layout: Layout,
    include_view_state: bool,
    include_previews: bool,
    shared_component: bool,
) -> TestResult<Vec<u8>> {
    let mut document = ArchiveObject::new(
        DOCUMENT_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-document".to_vec(),
            },
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: document_payload(
                    layout,
                    include_view_state.then_some(VIEW_STATE_IDENTIFIER),
                )?,
            },
            RawMessage {
                type_: 778,
                data: b"after-document".to_vec(),
            },
        ],
    )?;
    document.archive_info.message_infos[1].object_references =
        vec![BODY_IDENTIFIER, SECTION_IDENTIFIER];
    if include_view_state {
        document.archive_info.message_infos[1]
            .object_references
            .push(VIEW_STATE_IDENTIFIER);
        let mut field = FieldInfo::new(vec![15, 5]);
        field.object_references = vec![VIEW_STATE_IDENTIFIER];
        document.archive_info.message_infos[1]
            .field_infos
            .push(field);
    }

    let mut body = object(
        BODY_IDENTIFIER,
        STORAGE_MESSAGE_TYPE,
        tswp::StorageArchive {
            text: vec!["Semantic body text survives page layout edits".to_owned()],
            table_section: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(SECTION_IDENTIFIER)),
                }],
            }),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    body.archive_info.message_infos[0].object_references = vec![SECTION_IDENTIFIER];
    let section = object(
        SECTION_IDENTIFIER,
        SECTION_MESSAGE_TYPE,
        tp::SectionArchive {
            name: Some("Layout section".to_owned()),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;

    let objects = vec![document, body, section];
    let mut objects = objects;
    let mut view_component = None;
    if include_view_state {
        let mut bridge = object(
            VIEW_STATE_IDENTIFIER,
            VIEW_STATE_MESSAGE_TYPE,
            tsk::ViewStateArchive {
                view_state_root: reference(VIEW_STATE_ROOT_IDENTIFIER),
                ..Default::default()
            }
            .encode_to_vec(),
        )?;
        bridge.archive_info.message_infos[0].object_references = vec![VIEW_STATE_ROOT_IDENTIFIER];
        let mut bridge_field = FieldInfo::new(vec![1]);
        bridge_field.object_references = vec![VIEW_STATE_ROOT_IDENTIFIER];
        bridge.archive_info.message_infos[0]
            .field_infos
            .push(bridge_field);
        let mut view_payload = tp::ViewStateRootArchive {
            layout_state: Some(reference(LAYOUT_STATE_IDENTIFIER)),
            ..Default::default()
        }
        .encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(&mut view_payload, 99, 77_777)?;
        let mut view_state = ArchiveObject::new(
            VIEW_STATE_ROOT_IDENTIFIER,
            vec![
                RawMessage {
                    type_: 779,
                    data: b"before-view-state".to_vec(),
                },
                RawMessage {
                    type_: VIEW_STATE_ROOT_MESSAGE_TYPE,
                    data: view_payload,
                },
                RawMessage {
                    type_: 780,
                    data: b"after-view-state".to_vec(),
                },
            ],
        )?;
        let message_info = &mut view_state.archive_info.message_infos[1];
        message_info.object_references = vec![LAYOUT_STATE_IDENTIFIER];
        let mut field = FieldInfo::new(vec![1]);
        field.object_references = vec![LAYOUT_STATE_IDENTIFIER];
        message_info.field_infos.push(field);
        let mut detached_payload = tp::ViewStateRootArchive {
            layout_state: Some(reference(DETACHED_LAYOUT_STATE_IDENTIFIER)),
            ..Default::default()
        }
        .encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(&mut detached_payload, 98, 88_888)?;
        let mut detached = object(
            DETACHED_VIEW_STATE_ROOT_IDENTIFIER,
            VIEW_STATE_ROOT_MESSAGE_TYPE,
            detached_payload,
        )?;
        detached.archive_info.message_infos[0].object_references =
            vec![DETACHED_LAYOUT_STATE_IDENTIFIER];
        let mut detached_field = FieldInfo::new(vec![1]);
        detached_field.object_references = vec![DETACHED_LAYOUT_STATE_IDENTIFIER];
        detached.archive_info.message_infos[0]
            .field_infos
            .push(detached_field);
        let view_objects = vec![
            bridge,
            view_state,
            object(
                LAYOUT_STATE_IDENTIFIER,
                10_148,
                b"opaque layout-state cache".to_vec(),
            )?,
            detached,
            object(
                DETACHED_LAYOUT_STATE_IDENTIFIER,
                10_148,
                b"detached opaque layout-state cache".to_vec(),
            )?,
        ];
        if shared_component {
            objects.extend(view_objects);
        } else {
            view_component = Some(component(view_objects)?);
        }
    }

    let document = component_with_unknown_header(objects, DOCUMENT_IDENTIFIER)?;
    let unrelated = component(vec![object(900, 999, PRIVATE_MARKER.to_vec())?])?;
    let mut entries = vec![
        ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
        (DOCUMENT_MEMBER, document.as_slice()),
        (UNRELATED_MEMBER, unrelated.as_slice()),
    ];
    if let Some(view_component) = view_component.as_deref() {
        entries.push((VIEW_STATE_MEMBER, view_component));
    }
    if include_previews {
        entries.extend([
            (PREVIEW_NAMES[0], b"full preview".as_slice()),
            (PREVIEW_NAMES[1], b"micro preview".as_slice()),
            (PREVIEW_NAMES[2], b"web preview".as_slice()),
        ]);
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn legacy_package(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner = litchi_iwa_archive::package::to_bytes(inner_entries, Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.pages/Index.zip", inner.as_slice()),
            (
                "legacy.pages/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn component_stream(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other("missing synthetic Pages component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_document(
    package: &[u8],
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult,
) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, DOCUMENT_MEMBER)?;
    let mut archive = Archive::parse(&stream)?;
    let object = archive
        .object_mut(DOCUMENT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("missing synthetic Pages document"))?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing Pages document message"))?;
    mutate(object, message_index)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn rewrite_component(
    package: &[u8],
    component: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, component)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(
        catalog
            .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())?,
    )
}

fn with_overlong_object_length_prefix(
    package: &[u8],
    component: &str,
    identifier: u64,
) -> TestResult<Vec<u8>> {
    let mut stream = component_stream(package, component)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
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
    Ok(
        catalog
            .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())?,
    )
}

fn with_unknown_group(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |object, index| {
        let mut payload = object.messages[index].data.clone();
        let mut inner = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut inner, 1, 7)?;
        encode_varint_into(&mut payload, u64::from((99_u32 << 3) | 3));
        payload.extend_from_slice(&inner);
        encode_varint_into(&mut payload, u64::from((99_u32 << 3) | 4));
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn message_payload(package: &[u8], identifier: u64, message_type: u32) -> TestResult<Vec<u8>> {
    message_payload_in(package, DOCUMENT_MEMBER, identifier, message_type)
}

fn message_payload_in(
    package: &[u8],
    component: &str,
    identifier: u64,
    message_type: u32,
) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&component_stream(package, component)?)?;
    Ok(archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == message_type)
        })
        .ok_or_else(|| io::Error::other("missing synthetic Pages message"))?
        .data
        .clone())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, DOCUMENT_MEMBER)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let start = header_offset + prefix_length;
    let end = start + usize::try_from(length)?;
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

fn entry_names(package: &[u8]) -> TestResult<Vec<String>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .map(|entry| entry.name().to_owned())
        .collect())
}

fn assert_only_layout_components_and_previews_changed(before: &[u8], after: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(before)?;
    let after = Catalog::from_bytes(after)?;
    for old in before.iter() {
        if PREVIEW_NAMES.contains(&old.name()) {
            assert!(after.iter().all(|entry| entry.name() != old.name()));
            continue;
        }
        let new = after
            .iter()
            .find(|entry| entry.name() == old.name())
            .ok_or_else(|| io::Error::other("changed package lost an untouched entry"))?;
        if old.name() == DOCUMENT_MEMBER || old.name() == VIEW_STATE_MEMBER {
            assert_ne!(old.data(), new.data());
        } else {
            assert_eq!(old.data(), new.data());
            assert_eq!(old.metadata(), new.metadata());
            assert_eq!(
                old.raw_record().local_record(),
                new.raw_record().local_record()
            );
        }
    }
    Ok(())
}

fn changed_layout(before: Layout) -> TestResult<Layout> {
    let mut after = before;
    let orientation = match before.orientation() {
        Some(Orientation::Landscape) => Orientation::Portrait,
        Some(_) | None => Orientation::Landscape,
    };
    after.set_orientation(Some(orientation))?;
    Ok(after)
}

#[test]
fn native_fixture_noop_change_preview_deletion_semantics_and_inverse() -> TestResult {
    let bytes = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.page_layout()?;
    let text = package.text()?;
    let source_pointer = package.source_bytes().as_ptr();
    let preview_count = PREVIEW_NAMES
        .iter()
        .filter(|name| {
            entry_names(&bytes).is_ok_and(|names| names.iter().any(|entry| entry == **name))
        })
        .count();

    let mut noop = package.edit_page_layout()?;
    assert_eq!(noop.layout(), before);
    noop.set_layout(before)?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);
    assert_eq!(
        entry_names(noop.package().source_bytes())?,
        entry_names(&bytes)?
    );

    let after = changed_layout(before)?;
    let mut edit = package.edit_page_layout()?;
    edit.set_layout(after)?;
    let changed = edit.commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().page_layout()?, after);
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert_eq!(changed.diagnostics().deleted_previews(), preview_count);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_eq!(changed.package().text()?, text);
    for preview in PREVIEW_NAMES {
        assert!(
            entry_names(changed.package().source_bytes())?
                .iter()
                .all(|name| name != preview)
        );
    }
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let restored = changed
        .package()
        .apply_page_layout(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(restored.package().page_layout()?, before);
    assert_eq!(restored.package().text()?, text);
    Ok(())
}

#[test]
fn changed_layout_invalidates_view_state_and_preserves_unknowns_and_locality() -> TestResult {
    let before_layout = full_layout()?;
    let bytes = synthetic_package(before_layout, true, true)?;
    let package = Package::from_bytes(&bytes)?;
    let source_text = package.text()?;
    let source_document = message_payload(&bytes, DOCUMENT_IDENTIFIER, DOCUMENT_MESSAGE_TYPE)?;
    let source_header = object_header(&bytes, DOCUMENT_IDENTIFIER)?;
    let source_before = message_payload(&bytes, DOCUMENT_IDENTIFIER, 777)?;
    let source_after = message_payload(&bytes, DOCUMENT_IDENTIFIER, 778)?;
    let source_view = message_payload_in(
        &bytes,
        VIEW_STATE_MEMBER,
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    let source_view_before =
        message_payload_in(&bytes, VIEW_STATE_MEMBER, VIEW_STATE_ROOT_IDENTIFIER, 779)?;
    let source_view_after =
        message_payload_in(&bytes, VIEW_STATE_MEMBER, VIEW_STATE_ROOT_IDENTIFIER, 780)?;
    let source_layout_state =
        message_payload_in(&bytes, VIEW_STATE_MEMBER, LAYOUT_STATE_IDENTIFIER, 10_148)?;
    let source_detached = message_payload_in(
        &bytes,
        VIEW_STATE_MEMBER,
        DETACHED_VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    let source_detached_cache = message_payload_in(
        &bytes,
        VIEW_STATE_MEMBER,
        DETACHED_LAYOUT_STATE_IDENTIFIER,
        10_148,
    )?;

    let after_layout = changed_layout(before_layout)?;
    let mut edit = package.edit_page_layout()?;
    edit.set_layout(after_layout)?;
    let changed = edit.commit()?;
    assert_eq!(changed.package().page_layout()?, after_layout);
    assert_eq!(changed.package().text()?, source_text);
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert_eq!(changed.diagnostics().deleted_previews(), 3);
    assert_only_layout_components_and_previews_changed(&bytes, changed.package().source_bytes())?;
    assert_eq!(
        message_payload(changed.package().source_bytes(), DOCUMENT_IDENTIFIER, 777)?,
        source_before
    );
    assert_eq!(
        message_payload(changed.package().source_bytes(), DOCUMENT_IDENTIFIER, 778)?,
        source_after
    );
    let target_document = message_payload(
        changed.package().source_bytes(),
        DOCUMENT_IDENTIFIER,
        DOCUMENT_MESSAGE_TYPE,
    )?;
    assert_eq!(
        raw_fields(&target_document, 99)?,
        raw_fields(&source_document, 99)?
    );
    assert_eq!(
        raw_fields(
            &object_header(changed.package().source_bytes(), DOCUMENT_IDENTIFIER)?,
            99,
        )?,
        raw_fields(&source_header, 99)?
    );
    let target_view = message_payload_in(
        changed.package().source_bytes(),
        VIEW_STATE_MEMBER,
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    assert!(raw_fields(&target_view, 1)?.is_empty());
    assert_eq!(raw_fields(&target_view, 99)?, raw_fields(&source_view, 99)?);
    assert_eq!(
        message_payload_in(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            VIEW_STATE_ROOT_IDENTIFIER,
            779,
        )?,
        source_view_before
    );
    assert_eq!(
        message_payload_in(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            VIEW_STATE_ROOT_IDENTIFIER,
            780,
        )?,
        source_view_after
    );
    assert_eq!(
        message_payload_in(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            LAYOUT_STATE_IDENTIFIER,
            10_148,
        )?,
        source_layout_state
    );
    assert_eq!(
        message_payload_in(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            DETACHED_VIEW_STATE_ROOT_IDENTIFIER,
            VIEW_STATE_ROOT_MESSAGE_TYPE,
        )?,
        source_detached
    );
    assert_eq!(
        message_payload_in(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            DETACHED_LAYOUT_STATE_IDENTIFIER,
            10_148,
        )?,
        source_detached_cache
    );
    let changed_archive = Archive::parse(&component_stream(
        changed.package().source_bytes(),
        VIEW_STATE_MEMBER,
    )?)?;
    let view_object = changed_archive
        .object(VIEW_STATE_ROOT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("changed view-state root is missing"))?;
    let view_info = view_object
        .archive_info
        .message_infos
        .get(1)
        .ok_or_else(|| io::Error::other("changed view-state metadata is missing"))?;
    assert!(
        !view_info
            .object_references
            .contains(&LAYOUT_STATE_IDENTIFIER)
    );
    assert!(
        view_info
            .field_infos
            .iter()
            .all(|field| { !field.object_references.contains(&LAYOUT_STATE_IDENTIFIER) })
    );

    let restored = changed
        .package()
        .apply_page_layout(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    Ok(())
}

#[test]
fn shared_component_rooted_view_state_is_invalidated_but_detached_collision_is_preserved()
-> TestResult {
    let before = full_layout()?;
    let bytes = synthetic_package_with_view_location(before, true, false, true)?;
    let package = Package::from_bytes(&bytes)?;
    let detached = message_payload(
        &bytes,
        DETACHED_VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    let detached_cache = message_payload(&bytes, DETACHED_LAYOUT_STATE_IDENTIFIER, 10_148)?;

    let mut edit = package.edit_page_layout()?;
    edit.set_layout(changed_layout(before)?)?;
    let changed = edit.commit()?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    let rooted = message_payload(
        changed.package().source_bytes(),
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    assert!(raw_fields(&rooted, 1)?.is_empty());
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            DETACHED_VIEW_STATE_ROOT_IDENTIFIER,
            VIEW_STATE_ROOT_MESSAGE_TYPE,
        )?,
        detached
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            DETACHED_LAYOUT_STATE_IDENTIFIER,
            10_148,
        )?,
        detached_cache
    );
    let restored = changed
        .package()
        .apply_page_layout(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    Ok(())
}

#[test]
fn malformed_duplicate_and_wrong_wire_layout_fields_fail_closed() -> TestResult {
    let bytes = synthetic_package(full_layout()?, true, true)?;
    let duplicate = rewrite_document(&bytes, |object, index| {
        let field = WireView::parse(&object.messages[index].data)?
            .fields()
            .find(|field| field.number() == 30)
            .ok_or_else(|| io::Error::other("synthetic page width is missing"))?
            .raw()
            .to_vec();
        let mut payload = object.messages[index].data.clone();
        payload.extend_from_slice(&field);
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })?;
    let wrong_wire = rewrite_document(&bytes, |object, index| {
        let mut payload = object.messages[index].data.clone();
        litchi_iwa_common::wire::append_varint_field(&mut payload, 30, 612)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })?;
    let duplicate_message = rewrite_document(&bytes, |object, index| {
        object.push_message(object.messages[index].clone())?;
        Ok(())
    })?;

    for malformed in [duplicate, wrong_wire, duplicate_message] {
        match Package::from_bytes(&malformed) {
            Err(_ingress_error) => {},
            Ok(package) => {
                assert!(package.page_layout().is_err());
                assert_eq!(package.source_bytes(), malformed);
            },
        }
    }
    Ok(())
}

#[test]
fn rooted_view_state_header_reference_proof_fails_closed_at_every_hop() -> TestResult {
    let bytes = synthetic_package(full_layout()?, true, false)?;
    let missing_document_aggregate = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        object.archive_info.message_infos[1]
            .object_references
            .retain(|identifier| *identifier != VIEW_STATE_IDENTIFIER);
        Ok(())
    })?;
    let duplicate_bridge_aggregate = rewrite_component(&bytes, VIEW_STATE_MEMBER, |archive| {
        let object = archive
            .object_mut(VIEW_STATE_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing shared view-state object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(VIEW_STATE_ROOT_IDENTIFIER);
        Ok(())
    })?;
    let wrong_root_field_path = rewrite_component(&bytes, VIEW_STATE_MEMBER, |archive| {
        let object = archive
            .object_mut(VIEW_STATE_ROOT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing rooted view-state object"))?;
        let field = object.archive_info.message_infos[1]
            .field_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing rooted layout field metadata"))?;
        field.path = FieldPath::new(vec![2]);
        Ok(())
    })?;

    for malformed in [
        missing_document_aggregate,
        duplicate_bridge_aggregate,
        wrong_root_field_path,
    ] {
        let package = Package::from_bytes(&malformed)?;
        let before = package.page_layout()?;
        let mut edit = package.edit_page_layout()?;
        edit.set_layout(changed_layout(before)?)?;
        assert!(matches!(edit.commit(), Err(PageLayoutError::InvalidSource)));
        assert_eq!(package.source_bytes(), malformed);
    }
    Ok(())
}

#[test]
fn noncanonical_metadata_groups_and_deprecated_cache_edges_fail_changed_admission() -> TestResult {
    let bytes = synthetic_package(full_layout()?, false, false)?;
    let noncanonical =
        with_overlong_object_length_prefix(&bytes, DOCUMENT_MEMBER, SECTION_IDENTIFIER)?;
    let merge = rewrite_document(&bytes, |object, _index| {
        object.archive_info.should_merge = Some(true);
        Ok(())
    })?;
    let base = rewrite_document(&bytes, |object, index| {
        object.archive_info.message_infos[index].base_message_index = Some(0);
        Ok(())
    })?;
    let merge_version = rewrite_document(&bytes, |object, index| {
        object.archive_info.message_infos[index].diff_merge_version = vec![1];
        Ok(())
    })?;
    let diff_path = rewrite_document(&bytes, |object, index| {
        object.archive_info.message_infos[index].diff_field_path = Some(FieldPath::new(vec![30]));
        Ok(())
    })?;
    let remove = rewrite_document(&bytes, |object, index| {
        object.archive_info.message_infos[index]
            .fields_to_remove
            .push(FieldPath::new(vec![30]));
        Ok(())
    })?;
    let read_version = rewrite_document(&bytes, |object, index| {
        object.archive_info.message_infos[index].diff_read_version = vec![1];
        Ok(())
    })?;
    let unknown_group = with_unknown_group(&bytes)?;
    let deprecated_layout = rewrite_document(&bytes, |object, index| {
        let mut payload = object.messages[index].data.clone();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut payload,
            11,
            &reference(7_001).encode_to_vec(),
        )?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })?;
    let deprecated_view = rewrite_document(&bytes, |object, index| {
        let mut payload = object.messages[index].data.clone();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut payload,
            12,
            &reference(7_002).encode_to_vec(),
        )?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })?;

    let group_package = Package::from_bytes(&unknown_group)?;
    let group_layout = group_package.page_layout()?;
    let mut group_edit = group_package.edit_page_layout()?;
    group_edit.set_layout(changed_layout(group_layout)?)?;
    assert!(matches!(
        group_edit.commit(),
        Err(PageLayoutError::InvalidSource)
    ));
    assert_eq!(group_package.source_bytes(), unknown_group);

    for (label, adversarial) in [
        ("noncanonical", noncanonical),
        ("merge", merge),
        ("base", base),
        ("merge version", merge_version),
        ("diff path", diff_path),
        ("fields to remove", remove),
        ("read version", read_version),
        ("deprecated layout", deprecated_layout),
        ("deprecated view", deprecated_view),
    ] {
        let package = Package::from_bytes(&adversarial)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
        let before = package.page_layout()?;
        let mut edit = package.edit_page_layout()?;
        edit.set_layout(changed_layout(before)?)?;
        assert!(matches!(edit.commit(), Err(PageLayoutError::InvalidSource)));
        assert_eq!(package.source_bytes(), adversarial);
    }
    Ok(())
}

#[test]
fn legacy_nested_source_allows_noop_but_refuses_changed_layout() -> TestResult {
    let flat = synthetic_package(full_layout()?, false, false)?;
    let bytes = legacy_package(&flat)?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.page_layout()?;

    let mut noop = package.edit_page_layout()?;
    noop.set_layout(before)?;
    let noop = noop.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes(), bytes);

    let mut edit = package.edit_page_layout()?;
    edit.set_layout(changed_layout(before)?)?;
    assert!(matches!(
        edit.commit(),
        Err(PageLayoutError::UnsupportedSource)
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn stale_tampered_replayed_and_inverse_patches_conflict() -> TestResult {
    let bytes = synthetic_package(full_layout()?, false, false)?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.page_layout()?;
    let mut edit = package.edit_page_layout()?;
    edit.set_layout(changed_layout(before)?)?;
    let changed = edit.commit()?;

    assert!(matches!(
        changed.package().apply_page_layout(changed.patch()),
        Err(PageLayoutError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_page_layout(&changed.patch().inverse()),
        Err(PageLayoutError::PatchConflict)
    ));

    let catalog = Catalog::from_bytes(&bytes)?;
    let tampered_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered sentinel")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_bytes)?;
    assert_eq!(tampered.page_layout()?, before);
    assert!(matches!(
        tampered.apply_page_layout(changed.patch()),
        Err(PageLayoutError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn output_limit_is_typed_and_failure_atomic() -> TestResult {
    let bytes = synthetic_package(Layout::empty(), false, false)?;
    let unconstrained = Package::from_bytes(&bytes)?;
    let mut growth = unconstrained.edit_page_layout()?;
    growth.set_layout(full_layout()?)?;
    let growth = growth.commit()?;
    assert!(growth.package().source_bytes().len() > bytes.len());

    let limits = Limits::new(
        u64::try_from(bytes.len())?,
        32,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut edit = package.edit_page_layout()?;
    edit.set_layout(full_layout()?)?;
    assert!(matches!(
        edit.commit(),
        Err(PageLayoutError::LimitExceeded {
            kind: PageLayoutLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn public_transactions_are_send_sync_and_concurrent_reads_are_stable() -> TestResult {
    let bytes = synthetic_package(full_layout()?, true, true)?;
    let package = Arc::new(Package::from_bytes(&bytes)?);
    let expected = package.page_layout()?;
    let pointer = package.source_bytes().as_ptr();
    let handles = (0..8)
        .map(|_| {
            let package = Arc::clone(&package);
            std::thread::spawn(move || package.page_layout())
        })
        .collect::<Vec<_>>();
    for handle in handles {
        assert_eq!(
            handle
                .join()
                .map_err(|_| io::Error::other("page-layout reader panicked"))??,
            expected
        );
    }
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), pointer);
    assert_send_sync(package.as_ref());
    assert_type_send_sync::<Layout>();
    assert_type_send_sync::<PageLayoutEdit<'static>>();
    assert_type_send_sync::<PageLayoutCommit>();
    assert_type_send_sync::<PageLayoutPatch>();
    assert_type_send_sync::<PageLayoutDiagnostics>();
    assert_type_send_sync::<PageLayoutError>();
    assert_type_send_sync::<PageLayoutLimitKind>();
    Ok(())
}
