use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream, UnknownFieldRule,
};
use litchi_iwa_protos::{tp, tsa, tsk, tsp, tswp};
use litchi_pages::{
    Limits, Package,
    document_options::Options,
    document_settings::{
        Commit, Diagnostics, Edit, Error as SettingsError, LimitKind, Patch, Settings,
    },
    footnote::{self, Format, Gap, Kind, Numbering},
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SETTINGS_MEMBER: &str = "Index/Settings.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const DOCUMENT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const SECTION_IDENTIFIER: u64 = 43;
const VIEW_STATE_IDENTIFIER: u64 = 49;
const VIEW_STATE_ROOT_IDENTIFIER: u64 = 50;
const LAYOUT_STATE_IDENTIFIER: u64 = 51;
const SETTINGS_IDENTIFIER: u64 = 70;
const DETACHED_SETTINGS_IDENTIFIER: u64 = 80;
const OBSERVER_IDENTIFIER: u64 = 81;
const DETACHED_VIEW_ROOT_IDENTIFIER: u64 = 90;
const DETACHED_LAYOUT_STATE_IDENTIFIER: u64 = 91;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const SETTINGS_MESSAGE_TYPE: u32 = 10_012;
const SHARED_VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const VIEW_STATE_ROOT_MESSAGE_TYPE: u32 = 10_147;
const PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const PRIVATE_MARKER: &[u8] = b"private-pages-document-settings-marker-998244353";

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
    let (length, prefix_length) = decode_varint_from_bytes(&bytes[header_offset..])?;
    let start = header_offset
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header start overflow"))?;
    let end = start
        .checked_add(usize::try_from(length)?)
        .ok_or_else(|| io::Error::other("synthetic header end overflow"))?;
    assert_eq!(end, data_offset);
    let mut header = bytes[start..end].to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut rewritten = Vec::with_capacity(bytes.len().saturating_add(8));
    rewritten.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&bytes[data_offset..]);
    assert_eq!(Archive::parse(&rewritten)?.to_bytes()?, rewritten);
    Ok(SnappyStream::compress(&rewritten)?)
}

fn mixed_settings() -> TestResult<Settings> {
    Ok(Settings::new(
        Options::new(Some(false), Some(true), None, Some(false), None, Some(true)),
        footnote::Settings {
            kind: Some(Kind::Unknown(77)),
            format: None,
            numbering: Some(Numbering::RestartEachSection),
            gap: Some(Gap::new(37)?),
        },
    )?)
}

fn changed_both(before: Settings) -> TestResult<Settings> {
    let mut options = before.options();
    options.set_body_enabled(Some(true));
    options.set_headers_enabled(Some(false));
    options.set_footers_enabled(Some(false));
    options.set_facing_pages(Some(true));
    options.set_automatic_hyphenation(Some(true));
    options.set_ligatures_enabled(Some(false));
    let mut after = before;
    after.set_options(options);
    after.set_footnotes(footnote::Settings {
        kind: Some(Kind::SectionEndnotes),
        format: Some(Format::Unknown(81)),
        numbering: Some(Numbering::Unknown(82)),
        gap: Some(Gap::new(41)?),
    })?;
    Ok(after)
}

fn changed_options_only(before: Settings) -> Settings {
    let mut options = before.options();
    options.set_body_enabled(Some(!options.body_enabled().unwrap_or(true)));
    let mut after = before;
    after.set_options(options);
    after
}

fn changed_footnotes_only(before: Settings) -> TestResult<Settings> {
    let mut after = before;
    let old = before.footnotes();
    after.set_footnotes(footnote::Settings {
        kind: Some(if old.kind == Some(Kind::DocumentEndnotes) {
            Kind::SectionEndnotes
        } else {
            Kind::DocumentEndnotes
        }),
        format: Some(Format::Symbolic),
        numbering: Some(Numbering::RestartEachPage),
        gap: Some(Gap::new(
            old.gap.map(Gap::points).unwrap_or(0).saturating_add(1),
        )?),
    })?;
    Ok(after)
}

fn settings_payload(settings: Settings) -> TestResult<Vec<u8>> {
    let options = settings.options();
    let footnotes = settings.footnotes();
    let mut payload = tp::SettingsArchive {
        body: options.body_enabled(),
        headers: options.headers_enabled(),
        footers: options.footers_enabled(),
        facing_pages: options.facing_pages(),
        hyphenation: options.automatic_hyphenation(),
        use_ligatures: options.ligatures_enabled(),
        footnote_kind: footnotes.kind.map(Kind::as_raw),
        footnote_format: footnotes.format.map(Format::as_raw),
        footnote_numbering: footnotes.numbering.map(Numbering::as_raw),
        footnote_gap: footnotes
            .gap
            .map(Gap::points)
            .map(i32::try_from)
            .transpose()?,
        ..Default::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut payload, 99, 123_456)?;
    Ok(payload)
}

fn view_objects() -> TestResult<Vec<ArchiveObject>> {
    let mut bridge = object(
        VIEW_STATE_IDENTIFIER,
        SHARED_VIEW_STATE_MESSAGE_TYPE,
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

    let mut root_payload = tp::ViewStateRootArchive {
        layout_state: Some(reference(LAYOUT_STATE_IDENTIFIER)),
        ..Default::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut root_payload, 98, 8_888)?;
    let mut root = ArchiveObject::new(
        VIEW_STATE_ROOT_IDENTIFIER,
        vec![
            RawMessage {
                type_: 779,
                data: b"before-rooted-view-state".to_vec(),
            },
            RawMessage {
                type_: VIEW_STATE_ROOT_MESSAGE_TYPE,
                data: root_payload,
            },
            RawMessage {
                type_: 780,
                data: b"after-rooted-view-state".to_vec(),
            },
        ],
    )?;
    root.archive_info.message_infos[1].object_references = vec![LAYOUT_STATE_IDENTIFIER];
    let mut root_field = FieldInfo::new(vec![1]);
    root_field.object_references = vec![LAYOUT_STATE_IDENTIFIER];
    root.archive_info.message_infos[1]
        .field_infos
        .push(root_field);

    let mut detached_payload = tp::ViewStateRootArchive {
        layout_state: Some(reference(DETACHED_LAYOUT_STATE_IDENTIFIER)),
        ..Default::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut detached_payload, 97, 7_777)?;
    let mut detached = object(
        DETACHED_VIEW_ROOT_IDENTIFIER,
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

    Ok(vec![
        bridge,
        root,
        object(
            LAYOUT_STATE_IDENTIFIER,
            10_148,
            b"rooted layout cache".to_vec(),
        )?,
        detached,
        object(
            DETACHED_LAYOUT_STATE_IDENTIFIER,
            10_148,
            b"detached layout cache".to_vec(),
        )?,
    ])
}

fn synthetic_package(
    settings: Settings,
    split_components: bool,
    include_previews: bool,
) -> TestResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT_IDENTIFIER,
        DOCUMENT_MESSAGE_TYPE,
        tp::DocumentArchive {
            super_: tsa::DocumentArchive {
                super_: tsk::DocumentArchive::default(),
                view_state: Some(reference(VIEW_STATE_IDENTIFIER)),
                ..Default::default()
            },
            body_storage: Some(reference(BODY_IDENTIFIER)),
            section: Some(reference(SECTION_IDENTIFIER)),
            settings: Some(reference(SETTINGS_IDENTIFIER)),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let document_info = &mut document.archive_info.message_infos[0];
    document_info.object_references = vec![
        VIEW_STATE_IDENTIFIER,
        BODY_IDENTIFIER,
        SECTION_IDENTIFIER,
        SETTINGS_IDENTIFIER,
    ];
    for (path, identifier) in [
        (vec![15, 5], VIEW_STATE_IDENTIFIER),
        (vec![7], SETTINGS_IDENTIFIER),
    ] {
        let mut field = FieldInfo::new(path);
        field.object_references = vec![identifier];
        document_info.field_infos.push(field);
    }
    let mut body = object(
        BODY_IDENTIFIER,
        STORAGE_MESSAGE_TYPE,
        tswp::StorageArchive {
            text: vec!["Semantic text survives combined settings edits".to_owned()],
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
            name: Some("Settings section".to_owned()),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let selected_settings = ArchiveObject::new(
        SETTINGS_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-settings".to_vec(),
            },
            RawMessage {
                type_: SETTINGS_MESSAGE_TYPE,
                data: settings_payload(settings)?,
            },
            RawMessage {
                type_: 778,
                data: b"after-settings".to_vec(),
            },
        ],
    )?;
    let detached_settings = object(
        DETACHED_SETTINGS_IDENTIFIER,
        SETTINGS_MESSAGE_TYPE,
        settings_payload(changed_both(settings)?)?,
    )?;

    let mut document_objects = vec![document, body, section];
    let mut entries = Vec::<(&str, Vec<u8>)>::new();
    if split_components {
        entries.push((
            SETTINGS_MEMBER,
            component_with_unknown_header(
                vec![selected_settings, detached_settings],
                SETTINGS_IDENTIFIER,
            )?,
        ));
        entries.push((VIEW_STATE_MEMBER, component(view_objects()?)?));
    } else {
        document_objects.extend([selected_settings, detached_settings]);
        document_objects.extend(view_objects()?);
    }
    let document_component = if split_components {
        component(document_objects)?
    } else {
        component_with_unknown_header(document_objects, SETTINGS_IDENTIFIER)?
    };
    entries.push((DOCUMENT_MEMBER, document_component));
    entries.push((
        UNRELATED_MEMBER,
        component(vec![object(900, 999, PRIVATE_MARKER.to_vec())?])?,
    ));

    let mut package_entries = vec![("Data/sentinel.bin", b"ZIP sentinel".as_slice())];
    for (name, data) in &entries {
        package_entries.push((*name, data.as_slice()));
    }
    if include_previews {
        package_entries.extend([
            (PREVIEW_NAMES[0], b"full preview".as_slice()),
            (PREVIEW_NAMES[1], b"micro preview".as_slice()),
            (PREVIEW_NAMES[2], b"web preview".as_slice()),
        ]);
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        package_entries,
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

fn rewrite_settings_payload(
    package: &[u8],
    component: &str,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_component(package, component, |archive| {
        let object = archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing selected settings object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == SETTINGS_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing selected settings message"))?;
        let mut payload = object.messages[index].data.clone();
        mutate(&mut payload)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SETTINGS_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
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

fn message_payload(
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

fn object_header(package: &[u8], component: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, component)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
    let offset = usize::try_from(object.header_offset)?;
    let end = usize::try_from(object.data_offset)?;
    let (length, prefix) = decode_varint_from_bytes(&stream[offset..])?;
    let start = offset + prefix;
    assert_eq!(start + usize::try_from(length)?, end);
    Ok(stream[start..end].to_vec())
}

fn object_bytes(package: &[u8], component: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, component)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
    let start = usize::try_from(object.header_offset)?;
    let end = usize::try_from(
        object
            .data_offset
            .checked_add(object.data_length)
            .ok_or_else(|| io::Error::other("synthetic object end overflow"))?,
    )?;
    Ok(stream[start..end].to_vec())
}

fn object_location(package: &[u8], identifier: u64) -> TestResult<(String, ArchiveObject)> {
    let catalog = Catalog::from_bytes(package)?;
    let mut found = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?.into_bytes();
        let archive = Archive::parse(&stream)?;
        let Some(object) = archive.object(identifier) else {
            continue;
        };
        if found.is_some() {
            return Err(io::Error::other("duplicate synthetic object identifier").into());
        }
        found = Some((entry.name().to_owned(), object.clone()));
    }
    found.ok_or_else(|| io::Error::other("missing synthetic object identifier").into())
}

fn header_without_message_lengths(header: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let mut retained = Vec::new();
    for field in WireView::parse(header)?.fields() {
        if field.number() != 2 || field.wire_type() != 2 {
            retained.push(field.raw().to_vec());
            continue;
        }
        let mut message_info = vec![0x12];
        for nested in WireView::parse(field.payload())?.fields() {
            if nested.number() != 3 {
                message_info.extend_from_slice(nested.raw());
            }
        }
        retained.push(message_info);
    }
    Ok(retained)
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

fn assert_locality(before: &[u8], after: &[u8], changed_components: &[&str]) -> TestResult {
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
        if changed_components.contains(&old.name()) {
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

fn rewrite_document_payload(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_component(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing document message"))?;
        let mut payload = object.messages[index].data.clone();
        mutate(&mut payload)?;
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

fn assert_invalid_settings_source(bytes: &[u8]) -> TestResult {
    let package = Package::from_bytes(bytes)?;
    assert!(matches!(
        package.document_settings(),
        Err(SettingsError::InvalidSource)
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

fn assert_changed_settings_refused_atomically(bytes: &[u8], before: Settings) -> TestResult {
    let package = Package::from_bytes(bytes)?;
    assert_eq!(package.document_settings()?, before);
    assert!(matches!(
        package
            .edit_document_settings()?
            .set(changed_both(before)?)
            .commit(),
        Err(SettingsError::InvalidSource)
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn semantic_groups_presence_unknowns_gap_and_exact_noop_are_lossless() -> TestResult {
    let expected = mixed_settings()?;
    let bytes = synthetic_package(expected, true, true)?;
    let package = Package::from_bytes(&bytes)?;
    let pointer = package.source_bytes().as_ptr();
    let actual = package.document_settings()?;
    assert_eq!(actual, expected);
    assert_eq!(actual.options().body_enabled(), Some(false));
    assert_eq!(actual.options().headers_enabled(), Some(true));
    assert_eq!(actual.options().footers_enabled(), None);
    assert_eq!(actual.options().automatic_hyphenation(), None);
    assert_eq!(actual.footnotes().kind, Some(Kind::Unknown(77)));
    assert_eq!(actual.footnotes().format, None);
    assert_eq!(actual.footnotes().gap.map(Gap::points), Some(37));

    let commit = package.edit_document_settings()?.set(expected).commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.package().source_bytes().as_ptr(), pointer);
    assert_eq!(commit.package().source_bytes(), bytes);
    assert_eq!(
        entry_names(commit.package().source_bytes())?,
        entry_names(&bytes)?
    );
    let replay = package.apply_document_settings(commit.patch())?;
    assert!(replay.patch().is_noop());
    assert_eq!(replay.package().source_bytes().as_ptr(), pointer);
    Ok(())
}

#[test]
fn combined_change_is_atomic_preserves_locality_and_inverts_exactly() -> TestResult {
    let before = mixed_settings()?;
    let after = changed_both(before)?;
    let bytes = synthetic_package(before, true, true)?;
    let package = Package::from_bytes(&bytes)?;
    let text = package.text()?;
    let source_settings = message_payload(
        &bytes,
        SETTINGS_MEMBER,
        SETTINGS_IDENTIFIER,
        SETTINGS_MESSAGE_TYPE,
    )?;
    let source_header = object_header(&bytes, SETTINGS_MEMBER, SETTINGS_IDENTIFIER)?;
    let source_before = message_payload(&bytes, SETTINGS_MEMBER, SETTINGS_IDENTIFIER, 777)?;
    let source_after = message_payload(&bytes, SETTINGS_MEMBER, SETTINGS_IDENTIFIER, 778)?;
    let detached_settings = message_payload(
        &bytes,
        SETTINGS_MEMBER,
        DETACHED_SETTINGS_IDENTIFIER,
        SETTINGS_MESSAGE_TYPE,
    )?;
    let source_view = message_payload(
        &bytes,
        VIEW_STATE_MEMBER,
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    let source_view_before =
        message_payload(&bytes, VIEW_STATE_MEMBER, VIEW_STATE_ROOT_IDENTIFIER, 779)?;
    let source_view_after =
        message_payload(&bytes, VIEW_STATE_MEMBER, VIEW_STATE_ROOT_IDENTIFIER, 780)?;
    let source_cache = message_payload(&bytes, VIEW_STATE_MEMBER, LAYOUT_STATE_IDENTIFIER, 10_148)?;
    let detached_view = message_payload(
        &bytes,
        VIEW_STATE_MEMBER,
        DETACHED_VIEW_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    let detached_cache = message_payload(
        &bytes,
        VIEW_STATE_MEMBER,
        DETACHED_LAYOUT_STATE_IDENTIFIER,
        10_148,
    )?;

    let changed = package.edit_document_settings()?.set(after).commit()?;
    assert_eq!(changed.patch().before(), before);
    assert_eq!(changed.patch().after(), after);
    assert_eq!(changed.package().document_settings()?, after);
    assert_eq!(changed.package().text()?, text);
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert_eq!(changed.diagnostics().deleted_previews(), 3);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_locality(
        &bytes,
        changed.package().source_bytes(),
        &[SETTINGS_MEMBER, VIEW_STATE_MEMBER],
    )?;
    let target_settings = message_payload(
        changed.package().source_bytes(),
        SETTINGS_MEMBER,
        SETTINGS_IDENTIFIER,
        SETTINGS_MESSAGE_TYPE,
    )?;
    assert_eq!(
        raw_fields(&target_settings, 99)?,
        raw_fields(&source_settings, 99)?
    );
    assert_eq!(
        raw_fields(
            &object_header(
                changed.package().source_bytes(),
                SETTINGS_MEMBER,
                SETTINGS_IDENTIFIER,
            )?,
            99,
        )?,
        raw_fields(&source_header, 99)?
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            SETTINGS_MEMBER,
            SETTINGS_IDENTIFIER,
            777,
        )?,
        source_before
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            SETTINGS_MEMBER,
            SETTINGS_IDENTIFIER,
            778,
        )?,
        source_after
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            SETTINGS_MEMBER,
            DETACHED_SETTINGS_IDENTIFIER,
            SETTINGS_MESSAGE_TYPE,
        )?,
        detached_settings
    );
    let target_view = message_payload(
        changed.package().source_bytes(),
        VIEW_STATE_MEMBER,
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
    )?;
    assert!(raw_fields(&target_view, 1)?.is_empty());
    assert_eq!(raw_fields(&target_view, 98)?, raw_fields(&source_view, 98)?);
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            VIEW_STATE_ROOT_IDENTIFIER,
            779,
        )?,
        source_view_before
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            VIEW_STATE_ROOT_IDENTIFIER,
            780,
        )?,
        source_view_after
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            LAYOUT_STATE_IDENTIFIER,
            10_148,
        )?,
        source_cache
    );
    let changed_view_archive = Archive::parse(&component_stream(
        changed.package().source_bytes(),
        VIEW_STATE_MEMBER,
    )?)?;
    let changed_view_object = changed_view_archive
        .object(VIEW_STATE_ROOT_IDENTIFIER)
        .ok_or_else(|| io::Error::other("changed rooted view state is missing"))?;
    let changed_view_info = changed_view_object
        .archive_info
        .message_infos
        .get(1)
        .ok_or_else(|| io::Error::other("changed rooted view metadata is missing"))?;
    assert!(
        !changed_view_info
            .object_references
            .contains(&LAYOUT_STATE_IDENTIFIER)
    );
    assert!(
        changed_view_info
            .field_infos
            .iter()
            .all(|field| !field.object_references.contains(&LAYOUT_STATE_IDENTIFIER))
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            DETACHED_VIEW_ROOT_IDENTIFIER,
            VIEW_STATE_ROOT_MESSAGE_TYPE,
        )?,
        detached_view
    );
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            VIEW_STATE_MEMBER,
            DETACHED_LAYOUT_STATE_IDENTIFIER,
            10_148,
        )?,
        detached_cache
    );
    let applied = package.apply_document_settings(changed.patch())?;
    assert_eq!(
        applied.package().source_bytes(),
        changed.package().source_bytes()
    );
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.inverse(), changed.patch().clone());
    let restored = changed.package().apply_document_settings(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(restored.package().document_settings()?, before);
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn shared_component_combined_change_touches_one_component_and_restores() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, false, false)?;
    let package = Package::from_bytes(&bytes)?;
    let document_payload = message_payload(
        &bytes,
        DOCUMENT_MEMBER,
        DOCUMENT_IDENTIFIER,
        DOCUMENT_MESSAGE_TYPE,
    )?;
    let document_header = object_header(&bytes, DOCUMENT_MEMBER, DOCUMENT_IDENTIFIER)?;
    let changed = package
        .edit_document_settings()?
        .set(changed_both(before)?)
        .commit()?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_locality(&bytes, changed.package().source_bytes(), &[DOCUMENT_MEMBER])?;
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            DOCUMENT_MEMBER,
            DOCUMENT_IDENTIFIER,
            DOCUMENT_MESSAGE_TYPE,
        )?,
        document_payload
    );
    assert_eq!(
        object_header(
            changed.package().source_bytes(),
            DOCUMENT_MEMBER,
            DOCUMENT_IDENTIFIER,
        )?,
        document_header
    );
    let detached_settings = message_payload(
        &bytes,
        DOCUMENT_MEMBER,
        DETACHED_SETTINGS_IDENTIFIER,
        SETTINGS_MESSAGE_TYPE,
    )?;
    assert_eq!(
        message_payload(
            changed.package().source_bytes(),
            DOCUMENT_MEMBER,
            DETACHED_SETTINGS_IDENTIFIER,
            SETTINGS_MESSAGE_TYPE,
        )?,
        detached_settings
    );
    let restored = changed
        .package()
        .apply_document_settings(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    Ok(())
}

#[test]
fn root_reference_and_header_ownership_must_be_exact() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;

    let external = rewrite_document_payload(&bytes, |payload| {
        let mut document = tp::DocumentArchive::decode(payload.as_slice())?;
        document
            .settings
            .as_mut()
            .ok_or_else(|| io::Error::other("missing settings reference"))?
            .deprecated_is_external = Some(true);
        *payload = document.encode_to_vec();
        Ok(())
    })?;
    let zero = rewrite_document_payload(&bytes, |payload| {
        let mut document = tp::DocumentArchive::decode(payload.as_slice())?;
        document
            .settings
            .as_mut()
            .ok_or_else(|| io::Error::other("missing settings reference"))?
            .identifier = 0;
        *payload = document.encode_to_vec();
        Ok(())
    })?;
    let wrong = rewrite_document_payload(&bytes, |payload| {
        let mut document = tp::DocumentArchive::decode(payload.as_slice())?;
        document
            .settings
            .as_mut()
            .ok_or_else(|| io::Error::other("missing settings reference"))?
            .identifier = BODY_IDENTIFIER;
        *payload = document.encode_to_vec();
        Ok(())
    })?;
    let duplicate = rewrite_document_payload(&bytes, |payload| {
        let field = WireView::parse(payload)?
            .fields()
            .find(|field| field.number() == 7)
            .ok_or_else(|| io::Error::other("missing settings field"))?
            .raw()
            .to_vec();
        payload.extend_from_slice(&field);
        Ok(())
    })?;
    let missing_aggregate = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .retain(|identifier| *identifier != SETTINGS_IDENTIFIER);
        Ok(())
    })?;
    let duplicate_aggregate = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .push(SETTINGS_IDENTIFIER);
        Ok(())
    })?;
    let wrong_path = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        let field = object.archive_info.message_infos[0]
            .field_infos
            .iter_mut()
            .find(|field| field.object_references.contains(&SETTINGS_IDENTIFIER))
            .ok_or_else(|| io::Error::other("missing settings field metadata"))?;
        field.path = FieldPath::new(vec![8]);
        Ok(())
    })?;

    for malformed in [
        external,
        zero,
        wrong,
        duplicate,
        missing_aggregate,
        duplicate_aggregate,
        wrong_path,
    ] {
        assert_invalid_settings_source(&malformed)?;
    }

    let optional_path_absent = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing document object"))?;
        object.archive_info.message_infos[0]
            .field_infos
            .retain(|field| !field.object_references.contains(&SETTINGS_IDENTIFIER));
        Ok(())
    })?;
    let package = Package::from_bytes(&optional_path_absent)?;
    let changed = package
        .edit_document_settings()?
        .set(changed_both(before)?)
        .commit()?;
    assert_eq!(
        changed.package().document_settings()?,
        changed_both(before)?
    );
    assert_eq!(
        changed
            .package()
            .apply_document_settings(&changed.patch().inverse())?
            .package()
            .source_bytes(),
        optional_path_absent
    );
    Ok(())
}

#[test]
fn unrelated_observer_references_do_not_claim_settings_ownership_and_stay_exact() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;
    let bytes = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        let mut observer = object(
            OBSERVER_IDENTIFIER,
            55_555,
            b"detached observer payload".to_vec(),
        )?;
        let info = &mut observer.archive_info.message_infos[0];
        info.object_references = vec![SETTINGS_IDENTIFIER];
        let mut field = FieldInfo::new(vec![123, 456]);
        field.object_references = vec![SETTINGS_IDENTIFIER];
        info.field_infos.push(field);
        archive.objects.push(observer);
        Ok(())
    })?;
    let observer_bytes = object_bytes(&bytes, SETTINGS_MEMBER, OBSERVER_IDENTIFIER)?;
    let observer_header = object_header(&bytes, SETTINGS_MEMBER, OBSERVER_IDENTIFIER)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.document_settings()?, before);

    let changed = package
        .edit_document_settings()?
        .set(changed_both(before)?)
        .commit()?;
    assert_eq!(
        changed.package().document_settings()?,
        changed_both(before)?
    );
    assert_eq!(
        object_bytes(
            changed.package().source_bytes(),
            SETTINGS_MEMBER,
            OBSERVER_IDENTIFIER,
        )?,
        observer_bytes
    );
    assert_eq!(
        object_header(
            changed.package().source_bytes(),
            SETTINGS_MEMBER,
            OBSERVER_IDENTIFIER,
        )?,
        observer_header
    );
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn stale_rooted_cache_headers_and_no_layout_merge_diff_fail_atomically() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;

    let stale_document = rewrite_document_payload(&bytes, |payload| {
        let mut document = tp::DocumentArchive::decode(payload.as_slice())?;
        document.super_.view_state = None;
        *payload = document.encode_to_vec();
        Ok(())
    })?;
    let stale_view_root = rewrite_component(&bytes, VIEW_STATE_MEMBER, |archive| {
        let root = archive
            .object_mut(VIEW_STATE_ROOT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing rooted view state"))?;
        let index = root
            .messages
            .iter()
            .position(|message| message.type_ == VIEW_STATE_ROOT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing rooted view-state message"))?;
        let mut payload = tp::ViewStateRootArchive::decode(root.messages[index].data.as_slice())?;
        payload.layout_state = None;
        root.replace_message_preserving_header(
            index,
            RawMessage {
                type_: VIEW_STATE_ROOT_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let no_layout_merge = rewrite_component(&stale_view_root, VIEW_STATE_MEMBER, |archive| {
        archive
            .object_mut(VIEW_STATE_ROOT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing rooted view state"))?
            .archive_info
            .should_merge = Some(true);
        Ok(())
    })?;
    let no_layout_diff = rewrite_component(&stale_view_root, VIEW_STATE_MEMBER, |archive| {
        let root = archive
            .object_mut(VIEW_STATE_ROOT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing rooted view state"))?;
        let info = root
            .archive_info
            .message_infos
            .iter_mut()
            .find(|info| info.type_ == VIEW_STATE_ROOT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing rooted view-state metadata"))?;
        info.diff_field_path = Some(FieldPath::new(vec![1]));
        info.diff_read_version = vec![1];
        Ok(())
    })?;

    for adversarial in [
        stale_document,
        stale_view_root,
        no_layout_merge,
        no_layout_diff,
    ] {
        assert_changed_settings_refused_atomically(&adversarial, before)?;
    }
    Ok(())
}

#[test]
fn settings_payload_rejects_duplicate_wrong_wire_noncanonical_and_invalid_scalars() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;
    let duplicate = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        let raw = WireView::parse(payload)?
            .fields()
            .find(|field| field.number() == 1)
            .ok_or_else(|| io::Error::other("missing settings body field"))?
            .raw()
            .to_vec();
        payload.extend_from_slice(&raw);
        Ok(())
    })?;
    let wrong_wire = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        litchi_iwa_common::wire::append_length_delimited_field(payload, 3, b"false")?;
        Ok(())
    })?;
    let invalid_bool = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        litchi_iwa_common::wire::append_varint_field(payload, 3, 2)?;
        Ok(())
    })?;
    let noncanonical_bool = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        encode_varint_into(payload, u64::from(3_u32 << 3));
        payload.extend_from_slice(&[0x80, 0x00]);
        Ok(())
    })?;
    let noncanonical_int32 = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        encode_varint_into(payload, u64::from(31_u32 << 3));
        payload.extend_from_slice(&[0x81, 0x00]);
        Ok(())
    })?;

    for malformed in [
        duplicate,
        wrong_wire,
        invalid_bool,
        noncanonical_bool,
        noncanonical_int32,
    ] {
        assert_invalid_settings_source(&malformed)?;
    }
    Ok(())
}

#[test]
fn unknown_groups_are_read_but_changed_splice_fails_closed() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;
    let grouped = rewrite_settings_payload(&bytes, SETTINGS_MEMBER, |payload| {
        let mut inner = Vec::new();
        litchi_iwa_common::wire::append_varint_field(&mut inner, 1, 7)?;
        encode_varint_into(payload, u64::from((99_u32 << 3) | 3));
        payload.extend_from_slice(&inner);
        encode_varint_into(payload, u64::from((99_u32 << 3) | 4));
        Ok(())
    })?;
    let package = Package::from_bytes(&grouped)?;
    assert_eq!(package.document_settings()?, before);
    assert!(matches!(
        package
            .edit_document_settings()?
            .set(changed_both(before)?)
            .commit(),
        Err(SettingsError::InvalidSource)
    ));
    assert_eq!(package.source_bytes(), grouped);
    Ok(())
}

#[test]
fn noncanonical_prefix_and_all_merge_diff_metadata_are_refused_atomically() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;
    let prefix = with_overlong_object_length_prefix(&bytes, SETTINGS_MEMBER, SETTINGS_IDENTIFIER)?;
    let should_merge = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .should_merge = Some(true);
        Ok(())
    })?;
    let base = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .message_infos[1]
            .base_message_index = Some(0);
        Ok(())
    })?;
    let merge_version = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .message_infos[1]
            .diff_merge_version = vec![1];
        Ok(())
    })?;
    let diff_path = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .message_infos[1]
            .diff_field_path = Some(FieldPath::new(vec![1]));
        Ok(())
    })?;
    let fields_to_remove = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .message_infos[1]
            .fields_to_remove
            .push(FieldPath::new(vec![1]));
        Ok(())
    })?;
    let read_version = rewrite_component(&bytes, SETTINGS_MEMBER, |archive| {
        archive
            .object_mut(SETTINGS_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing settings object"))?
            .archive_info
            .message_infos[1]
            .diff_read_version = vec![1];
        Ok(())
    })?;

    let prefix_package = Package::from_bytes(&prefix)?;
    assert_eq!(prefix_package.document_settings()?, before);
    assert!(matches!(
        prefix_package
            .edit_document_settings()?
            .set(changed_both(before)?)
            .commit(),
        Err(SettingsError::InvalidSource)
    ));
    assert_eq!(prefix_package.source_bytes(), prefix);

    for malformed in [
        should_merge,
        base,
        merge_version,
        diff_path,
        fields_to_remove,
        read_version,
    ] {
        assert_invalid_settings_source(&malformed)?;
    }
    Ok(())
}

#[test]
fn legacy_noop_is_exact_but_changed_edit_is_unsupported() -> TestResult {
    let before = mixed_settings()?;
    let flat = synthetic_package(before, false, false)?;
    let bytes = legacy_package(&flat)?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.document_settings()?, before);
    let pointer = package.source_bytes().as_ptr();
    let noop = package.edit_document_settings()?.set(before).commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes().as_ptr(), pointer);
    assert_eq!(noop.package().source_bytes(), bytes);
    assert!(matches!(
        package
            .edit_document_settings()?
            .set(changed_both(before)?)
            .commit(),
        Err(SettingsError::UnsupportedSource)
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn patches_reject_replay_inverse_on_source_tamper_and_competing_targets() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, false)?;
    let package = Package::from_bytes(&bytes)?;
    let first = package
        .edit_document_settings()?
        .set(changed_options_only(before))
        .commit()?;
    let second = package
        .edit_document_settings()?
        .set(changed_footnotes_only(before)?)
        .commit()?;

    assert!(matches!(
        first.package().apply_document_settings(first.patch()),
        Err(SettingsError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_document_settings(&first.patch().inverse()),
        Err(SettingsError::PatchConflict)
    ));
    assert!(matches!(
        first.package().apply_document_settings(second.patch()),
        Err(SettingsError::PatchConflict)
    ));

    let catalog = Catalog::from_bytes(&bytes)?;
    let tampered_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"tampered ZIP sentinel",
        )],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_bytes)?;
    assert_eq!(tampered.document_settings()?, before);
    assert!(matches!(
        tampered.apply_document_settings(first.patch()),
        Err(SettingsError::PatchConflict)
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn output_limit_is_typed_and_failure_is_atomic() -> TestResult {
    let before = Settings::default();
    let after = changed_both(before)?;
    let bytes = rewrite_component(
        &synthetic_package(before, false, false)?,
        DOCUMENT_MEMBER,
        |archive| {
            let document = archive
                .object_mut(DOCUMENT_IDENTIFIER)
                .ok_or_else(|| io::Error::other("missing document object"))?;
            let message = document
                .messages
                .iter()
                .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
                .ok_or_else(|| io::Error::other("missing document message"))?;
            let mut payload =
                tp::DocumentArchive::decode(document.messages[message].data.as_slice())?;
            payload.super_.view_state = None;
            document.replace_message_preserving_header(
                message,
                RawMessage {
                    type_: DOCUMENT_MESSAGE_TYPE,
                    data: payload.encode_to_vec(),
                },
            )?;
            let info = &mut document.archive_info.message_infos[message];
            info.object_references
                .retain(|identifier| *identifier != VIEW_STATE_IDENTIFIER);
            info.field_infos
                .retain(|field| !field.object_references.contains(&VIEW_STATE_IDENTIFIER));
            Ok(())
        },
    )?;
    let unconstrained = Package::from_bytes(&bytes)?;
    let grown = unconstrained
        .edit_document_settings()?
        .set(after)
        .commit()?;
    assert!(grown.package().source_bytes().len() > bytes.len());

    let limits = Limits::new(
        u64::try_from(bytes.len())?,
        32,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    assert!(matches!(
        package.edit_document_settings()?.set(after).commit(),
        Err(SettingsError::LimitExceeded {
            kind: LimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn native_basic_pages_combined_change_preserves_text_and_inverts_exactly() -> TestResult {
    let bytes = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.document_settings()?;
    let text = package.text()?;
    let (_document_component, document_object) = object_location(&bytes, DOCUMENT_IDENTIFIER)?;
    let document_message = document_object
        .messages
        .iter()
        .find(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native document message is missing"))?;
    let document = tp::DocumentArchive::decode(document_message.data.as_slice())?;
    let settings_identifier = document
        .settings
        .ok_or_else(|| io::Error::other("native settings reference is missing"))?
        .identifier;
    let view_identifier = document
        .super_
        .view_state
        .ok_or_else(|| io::Error::other("native view-state reference is missing"))?
        .identifier;
    let (settings_component, settings_object) = object_location(&bytes, settings_identifier)?;
    let settings_index = settings_object
        .messages
        .iter()
        .position(|message| message.type_ == SETTINGS_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native settings message is missing"))?;
    let settings_info = settings_object.archive_info.message_infos[settings_index].clone();
    assert_eq!(settings_info.versions, [1, 0, 5]);
    let gap_info = settings_info
        .field_infos
        .iter()
        .find(|field| field.path.as_slice() == [34])
        .ok_or_else(|| io::Error::other("native facing-pages metadata is missing"))?;
    assert_eq!(
        gap_info.effective_unknown_field_rule(),
        UnknownFieldRule::IgnoreAndPreserve
    );
    let settings_header = object_header(&bytes, &settings_component, settings_identifier)?;

    let (_bridge_component, bridge_object) = object_location(&bytes, view_identifier)?;
    let bridge_message = bridge_object
        .messages
        .iter()
        .find(|message| message.type_ == SHARED_VIEW_STATE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native view-state bridge is missing"))?;
    let view_root_identifier = tsk::ViewStateArchive::decode(bridge_message.data.as_slice())?
        .view_state_root
        .identifier;
    let (view_component, view_root) = object_location(&bytes, view_root_identifier)?;
    let view_index = view_root
        .messages
        .iter()
        .position(|message| message.type_ == VIEW_STATE_ROOT_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("native rooted view-state message is missing"))?;
    let view_payload =
        tp::ViewStateRootArchive::decode(view_root.messages[view_index].data.as_slice())?;
    let layout_identifier = view_payload
        .layout_state
        .ok_or_else(|| io::Error::other("native layout-state reference is missing"))?
        .identifier;
    let source_view_info = view_root.archive_info.message_infos[view_index].clone();
    let preview_count = Catalog::from_bytes(&bytes)?
        .iter()
        .filter(|entry| PREVIEW_NAMES.contains(&entry.name()))
        .count();
    let after = changed_both(before)?;
    let changed = package.edit_document_settings()?.set(after).commit()?;
    assert_eq!(changed.package().document_settings()?, after);
    assert_eq!(changed.package().text()?, text);
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 2);
    assert_eq!(changed.diagnostics().deleted_previews(), preview_count);
    assert!(changed.diagnostics().full_reparse_performed());

    let (target_settings_component, target_settings_object) =
        object_location(changed.package().source_bytes(), settings_identifier)?;
    assert_eq!(target_settings_component, settings_component);
    let target_settings_index = target_settings_object
        .messages
        .iter()
        .position(|message| message.type_ == SETTINGS_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("changed native settings message is missing"))?;
    let target_settings_info =
        &target_settings_object.archive_info.message_infos[target_settings_index];
    let mut normalized_target_settings = target_settings_info.clone();
    normalized_target_settings.length = settings_info.length;
    assert_eq!(normalized_target_settings, settings_info);
    assert_eq!(
        header_without_message_lengths(&object_header(
            changed.package().source_bytes(),
            &settings_component,
            settings_identifier,
        )?)?,
        header_without_message_lengths(&settings_header)?
    );

    let (target_view_component, target_view_root) =
        object_location(changed.package().source_bytes(), view_root_identifier)?;
    assert_eq!(target_view_component, view_component);
    let target_view_index = target_view_root
        .messages
        .iter()
        .position(|message| message.type_ == VIEW_STATE_ROOT_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("changed native rooted view state is missing"))?;
    let target_view_info = &target_view_root.archive_info.message_infos[target_view_index];
    let mut expected_view_info = source_view_info;
    expected_view_info.length = target_view_info.length;
    expected_view_info
        .object_references
        .retain(|identifier| *identifier != layout_identifier);
    for field in &mut expected_view_info.field_infos {
        field
            .object_references
            .retain(|identifier| *identifier != layout_identifier);
    }
    assert_eq!(*target_view_info, expected_view_info);
    for preview in PREVIEW_NAMES {
        assert!(
            Catalog::from_bytes(changed.package().source_bytes())?
                .iter()
                .all(|entry| entry.name() != preview)
        );
    }
    let restored = changed
        .package()
        .apply_document_settings(&changed.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(restored.package().document_settings()?, before);
    assert_eq!(restored.package().text()?, text);
    Ok(())
}

#[test]
fn transactions_are_redacted_send_sync_and_concurrent_changes_are_independent() -> TestResult {
    let before = mixed_settings()?;
    let bytes = synthetic_package(before, true, true)?;
    let package = Arc::new(Package::from_bytes(&bytes)?);
    let pointer = package.source_bytes().as_ptr();
    let options_package = Arc::clone(&package);
    let footnotes_package = Arc::clone(&package);
    let options = std::thread::spawn(move || {
        options_package
            .edit_document_settings()?
            .set(changed_options_only(before))
            .commit()
    });
    let footnotes = std::thread::spawn(move || {
        footnotes_package
            .edit_document_settings()?
            .set(changed_footnotes_only(before).map_err(|_| SettingsError::InvalidSource)?)
            .commit()
    });
    let options = options
        .join()
        .map_err(|_| io::Error::other("options editor panicked"))??;
    let footnotes = footnotes
        .join()
        .map_err(|_| io::Error::other("footnotes editor panicked"))??;
    assert_eq!(
        options.package().document_settings()?.footnotes(),
        before.footnotes()
    );
    assert_eq!(
        footnotes.package().document_settings()?.options(),
        before.options()
    );
    assert_ne!(
        options.package().source_bytes(),
        footnotes.package().source_bytes()
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), pointer);

    let edit_debug = format!("{:?}", package.edit_document_settings()?);
    let patch_debug = format!("{:?}", options.patch());
    let error_debug = format!("{:?}", SettingsError::InvalidSource);
    for rendered in [&edit_debug, &patch_debug, &error_debug] {
        assert!(!rendered.contains(DOCUMENT_MEMBER));
        assert!(!rendered.contains(UNRELATED_MEMBER));
        assert!(!rendered.contains(std::str::from_utf8(PRIVATE_MARKER)?));
    }

    assert_send_sync(package.as_ref());
    assert_type_send_sync::<Settings>();
    assert_type_send_sync::<Edit<'static>>();
    assert_type_send_sync::<Commit>();
    assert_type_send_sync::<Patch>();
    assert_type_send_sync::<Diagnostics>();
    assert_type_send_sync::<SettingsError>();
    assert_type_send_sync::<LimitKind>();
    Ok(())
}
