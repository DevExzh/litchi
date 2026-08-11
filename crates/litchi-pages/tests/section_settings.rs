use std::collections::BTreeMap;
use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsd, tsk, tsp, tswp};
use litchi_pages::{
    Package, Position, SectionSelector,
    section::{
        PageNumber, PageNumbering, Settings, Start,
        settings::{Commit, DependencyKind, Diagnostics, Edit, Error, LimitKind, Patch, Path},
    },
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const DOCUMENT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const VIEW_STATE_IDENTIFIER: u64 = 49;
const VIEW_STATE_ROOT_IDENTIFIER: u64 = 50;
const LAYOUT_STATE_IDENTIFIER: u64 = 51;
const FIRST_TEMPLATES: [u64; 3] = [60, 61, 62];
const SECOND_TEMPLATES: [u64; 3] = [63, 64, 65];
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const VIEW_STATE_MESSAGE_TYPE: u32 = 210;
const VIEW_STATE_ROOT_MESSAGE_TYPE: u32 = 10_147;
const LAYOUT_STATE_MESSAGE_TYPE: u32 = 10_148;
const PRIVATE_MARKER: &str = "private-pages-section-settings-marker-998244353";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const SETTINGS_FIELDS: [u32; 8] = [17, 18, 19, 20, 21, 22, 26, 28];

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/basic.pages")
}

fn assert_type_send_sync<T: Send + Sync>() {}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
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
    append_varint_field(&mut header, 99, 9_999)?;

    let mut with_unknown = Vec::with_capacity(bytes.len().saturating_add(8));
    with_unknown.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut with_unknown, u64::try_from(header.len())?);
    with_unknown.extend_from_slice(&header);
    with_unknown.extend_from_slice(&bytes[data_offset..]);

    let reparsed = Archive::parse(&with_unknown)?;
    assert_eq!(reparsed.to_bytes()?, with_unknown);
    Ok(SnappyStream::compress(&with_unknown)?)
}

fn settings(
    name: Option<&str>,
    booleans: [Option<bool>; 4],
    start: Option<Start>,
    numbering: Option<PageNumbering>,
    page: Option<u32>,
) -> TestResult<Settings> {
    let mut value = Settings::new();
    value.set_name(name)?;
    value.set_inherit_previous_header_footer(booleans[0]);
    value.set_first_page_different(booleans[1]);
    value.set_even_odd_pages_different(booleans[2]);
    value.set_first_page_hides_header_footer(booleans[3]);
    value.set_start(start)?;
    value.set_page_numbering(numbering)?;
    value.set_starting_page_number(page.map(PageNumber::new).transpose()?);
    Ok(value)
}

fn fixture_settings() -> TestResult<Settings> {
    settings(
        Some("Blank"),
        [Some(true), Some(false), Some(false), Some(false)],
        Some(Start::NextPage),
        Some(PageNumbering::ContinueFromPrevious),
        Some(1),
    )
}

fn changed_settings() -> TestResult<Settings> {
    settings(
        Some("Private aggregate chapter 東京"),
        [Some(true), Some(true), Some(true), Some(true)],
        Some(Start::LeftPage),
        Some(PageNumbering::Restart),
        Some(42),
    )
}

fn section_payload(
    value: &Settings,
    templates: [u64; 3],
    unknown_value: u64,
) -> TestResult<Vec<u8>> {
    let mut data = tp::SectionArchive {
        inherit_previous_header_footer: value.inherit_previous_header_footer(),
        section_template_first_page_different: value.first_page_different(),
        section_template_even_odd_pages_different: value.even_odd_pages_different(),
        section_start_kind: value.start().map(Start::as_raw),
        section_page_number_kind: value.page_numbering().map(PageNumbering::as_raw),
        section_page_number_start: value.starting_page_number().map(PageNumber::get),
        first_section_template_page: Some(reference(templates[0])),
        even_section_template_page: Some(reference(templates[1])),
        odd_section_template_page: Some(reference(templates[2])),
        name: value.name().map(str::to_owned),
        section_template_first_page_hides_header_footer: value.first_page_hides_header_footer(),
        background_fill: Some(tsd::FillArchive::default()),
        ..tp::SectionArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 99, unknown_value)?;
    Ok(data)
}

fn section_message(
    value: &Settings,
    templates: [u64; 3],
    unknown_value: u64,
) -> TestResult<RawMessage> {
    Ok(RawMessage {
        type_: SECTION_MESSAGE_TYPE,
        data: section_payload(value, templates, unknown_value)?,
    })
}

fn synthetic_package(first: &Settings, second: &Settings) -> TestResult<Vec<u8>> {
    synthetic_package_with_first_message(section_message(first, FIRST_TEMPLATES, 7_777)?, second)
}

fn synthetic_package_with_first_message(
    first_message: RawMessage,
    second: &Settings,
) -> TestResult<Vec<u8>> {
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
    let mut first_section = ArchiveObject::new(
        FIRST_SECTION_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-section-payload".to_vec(),
            },
            first_message,
            RawMessage {
                type_: 778,
                data: b"after-section-payload".to_vec(),
            },
        ],
    )?;
    first_section.archive_info.message_infos[1].object_references = FIRST_TEMPLATES.to_vec();
    for (field_number, identifier) in [23_u32, 24, 25].into_iter().zip(FIRST_TEMPLATES) {
        let mut field = FieldInfo::new(vec![field_number]);
        field.object_references = vec![identifier];
        first_section.archive_info.message_infos[1]
            .field_infos
            .push(field);
    }
    let mut second_section = ArchiveObject::new(
        SECOND_SECTION_IDENTIFIER,
        vec![section_message(second, SECOND_TEMPLATES, 8_888)?],
    )?;
    second_section.archive_info.message_infos[0].object_references = SECOND_TEMPLATES.to_vec();
    for (field_number, identifier) in [23_u32, 24, 25].into_iter().zip(SECOND_TEMPLATES) {
        let mut field = FieldInfo::new(vec![field_number]);
        field.object_references = vec![identifier];
        second_section.archive_info.message_infos[0]
            .field_infos
            .push(field);
    }
    let mut document_root = object(DOCUMENT_IDENTIFIER, 10_000, root.encode_to_vec())?;
    document_root.archive_info.message_infos[0].object_references =
        vec![BODY_IDENTIFIER, VIEW_STATE_IDENTIFIER];
    let mut body_field = FieldInfo::new(vec![4]);
    body_field.object_references = vec![BODY_IDENTIFIER];
    let mut view_state_field = FieldInfo::new(vec![15, 5]);
    view_state_field.object_references = vec![VIEW_STATE_IDENTIFIER];
    document_root.archive_info.message_infos[0]
        .field_infos
        .extend([body_field, view_state_field]);
    let mut body_object = object(BODY_IDENTIFIER, 2_001, body.encode_to_vec())?;
    body_object.archive_info.message_infos[0].object_references =
        vec![FIRST_SECTION_IDENTIFIER, SECOND_SECTION_IDENTIFIER];
    let document = component_with_unknown_header(
        vec![
            document_root,
            body_object,
            first_section,
            second_section,
            object(
                FIRST_TEMPLATES[0],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
            object(
                FIRST_TEMPLATES[1],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
            object(
                FIRST_TEMPLATES[2],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
            object(
                SECOND_TEMPLATES[0],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
            object(
                SECOND_TEMPLATES[1],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
            object(
                SECOND_TEMPLATES[2],
                10_143,
                tp::SectionTemplateArchive::default().encode_to_vec(),
            )?,
        ],
        FIRST_SECTION_IDENTIFIER,
    )?;
    let mut view_bridge = object(
        VIEW_STATE_IDENTIFIER,
        VIEW_STATE_MESSAGE_TYPE,
        tsk::ViewStateArchive {
            view_state_root: reference(VIEW_STATE_ROOT_IDENTIFIER),
            ..tsk::ViewStateArchive::default()
        }
        .encode_to_vec(),
    )?;
    view_bridge.archive_info.message_infos[0].object_references = vec![VIEW_STATE_ROOT_IDENTIFIER];
    let mut bridge_field = FieldInfo::new(vec![1]);
    bridge_field.object_references = vec![VIEW_STATE_ROOT_IDENTIFIER];
    view_bridge.archive_info.message_infos[0]
        .field_infos
        .push(bridge_field);
    let mut view_root_payload = tp::ViewStateRootArchive {
        layout_state: Some(reference(LAYOUT_STATE_IDENTIFIER)),
        ..tp::ViewStateRootArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut view_root_payload, 99, 99_999)?;
    let mut view_root = object(
        VIEW_STATE_ROOT_IDENTIFIER,
        VIEW_STATE_ROOT_MESSAGE_TYPE,
        view_root_payload,
    )?;
    view_root.archive_info.message_infos[0].object_references = vec![LAYOUT_STATE_IDENTIFIER];
    let mut layout_field = FieldInfo::new(vec![1]);
    layout_field.object_references = vec![LAYOUT_STATE_IDENTIFIER];
    view_root.archive_info.message_infos[0]
        .field_infos
        .push(layout_field);
    let view_state = component(vec![
        view_bridge,
        view_root,
        object(
            LAYOUT_STATE_IDENTIFIER,
            LAYOUT_STATE_MESSAGE_TYPE,
            b"rooted opaque layout cache".to_vec(),
        )?,
    ])?;
    let unrelated = component(vec![object(99, 777, PRIVATE_MARKER.as_bytes().to_vec())?])?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document.as_slice()),
            (VIEW_STATE_MEMBER, view_state.as_slice()),
            (UNRELATED_MEMBER, unrelated.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
        litchi_pages::Limits::default(),
    )?)
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner =
        litchi_iwa_archive::package::to_bytes(inner_entries, litchi_pages::Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.pages/Index.zip", inner.as_slice()),
            (
                "legacy.pages/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        litchi_pages::Limits::default(),
    )?)
}

fn member_stream(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing synthetic Pages component member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    member_stream(package, DOCUMENT_MEMBER)
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    message_payload_in(package, DOCUMENT_MEMBER, identifier, type_)
}

fn message_payload_in(
    package: &[u8],
    member: &str,
    identifier: u64,
    type_: u32,
) -> TestResult<Vec<u8>> {
    let stream = member_stream(package, member)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic Pages message"))?;
    Ok(message.data.clone())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic Pages object"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length_u64, prefix_length) = decode_varint_from_bytes(
        stream
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
    Ok(stream[header_start..header_end].to_vec())
}

fn normalize_message_length(header: &[u8], target_index: usize) -> TestResult<Vec<u8>> {
    let view = WireView::parse(header)?;
    let mut output = Vec::with_capacity(header.len());
    let mut message_index = 0usize;
    for field in view.fields() {
        if field.number() != 2 || field.wire_type() != 2 {
            output.extend_from_slice(field.raw());
            continue;
        }
        if message_index != target_index {
            output.extend_from_slice(field.raw());
            message_index += 1;
            continue;
        }
        output.extend_from_slice(field.key());
        output.extend_from_slice(b"<message-info>");
        let message = WireView::parse(field.payload())?;
        let effective_length = message
            .fields()
            .enumerate()
            .filter_map(|(index, nested)| {
                (nested.number() == 3 && nested.wire_type() == 0).then_some(index)
            })
            .last()
            .ok_or_else(|| io::Error::other("synthetic MessageInfo has no length"))?;
        for (index, nested) in message.fields().enumerate() {
            if index == effective_length {
                output.extend_from_slice(nested.key());
                output.extend_from_slice(b"<payload-length>");
            } else {
                output.extend_from_slice(nested.raw());
            }
        }
        output.extend_from_slice(b"</message-info>");
        message_index += 1;
    }
    Ok(output)
}

fn fields_except_settings(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse_with_limits(payload, WireLimits::default())?
        .fields()
        .filter(|field| !SETTINGS_FIELDS.contains(&field.number()))
        .map(|field| field.raw().to_vec())
        .collect())
}

fn entry_map(package: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

fn assert_previews(package: &[u8], expected: bool) -> TestResult {
    let entries = entry_map(package)?;
    for preview in PREVIEWS {
        assert_eq!(entries.contains_key(preview), expected, "preview {preview}");
    }
    Ok(())
}

fn assert_one_iwa_member_changed(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    assert_eq!(
        before.iter().map(|entry| entry.name()).collect::<Vec<_>>(),
        after.iter().map(|entry| entry.name()).collect::<Vec<_>>(),
    );
    let mut changed = Vec::new();
    for old in before.iter() {
        let new = after
            .iter()
            .find(|entry| entry.name() == old.name())
            .ok_or_else(|| io::Error::other("changed package lost a source entry"))?;
        if old.data() != new.data() {
            changed.push(old.name().to_owned());
            continue;
        }
        assert_eq!(
            old.metadata(),
            new.metadata(),
            "metadata changed: {}",
            old.name()
        );
        assert_eq!(
            old.raw_record().local_record(),
            new.raw_record().local_record(),
            "local record changed: {}",
            old.name()
        );
    }
    assert_eq!(changed.len(), 1, "unexpected changed members: {changed:?}");
    assert!(changed[0].ends_with(".iwa"), "non-IWA member changed");
    for preview in PREVIEWS {
        assert!(!changed.iter().any(|name| name == preview));
    }
    Ok(())
}

fn assert_changed_locality(source: &[u8], target: &[u8]) -> TestResult {
    assert_one_iwa_member_changed(source, target)?;
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let before_document = before
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("source document member is missing"))?;
    let after_document = after
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("target document member is missing"))?;
    assert_ne!(before_document.data(), after_document.data());
    for old in before.iter() {
        if old.name() == DOCUMENT_MEMBER {
            continue;
        }
        let new = after
            .iter()
            .find(|entry| entry.name() == old.name())
            .ok_or_else(|| io::Error::other("changed package lost an untouched entry"))?;
        assert_eq!(
            old.data(),
            new.data(),
            "member bytes changed: {}",
            old.name()
        );
        assert_eq!(
            old.metadata(),
            new.metadata(),
            "member metadata changed: {}",
            old.name()
        );
        assert_eq!(
            old.raw_record().local_record(),
            new.raw_record().local_record(),
            "member local record changed: {}",
            old.name()
        );
    }
    assert_previews(target, true)?;

    let source_first = message_payload(source, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let target_first = message_payload(target, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    assert_eq!(
        fields_except_settings(&source_first)?,
        fields_except_settings(&target_first)?,
    );
    assert_eq!(
        message_payload(source, SECOND_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?,
        message_payload(target, SECOND_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?,
    );
    assert_eq!(
        message_payload(source, BODY_IDENTIFIER, 2_001)?,
        message_payload(target, BODY_IDENTIFIER, 2_001)?,
    );
    assert_eq!(
        message_payload(source, DOCUMENT_IDENTIFIER, 10_000)?,
        message_payload(target, DOCUMENT_IDENTIFIER, 10_000)?,
    );
    for identifier in [
        DOCUMENT_IDENTIFIER,
        BODY_IDENTIFIER,
        SECOND_SECTION_IDENTIFIER,
    ] {
        assert_eq!(
            object_header(source, identifier)?,
            object_header(target, identifier)?,
        );
    }
    for identifier in FIRST_TEMPLATES.into_iter().chain(SECOND_TEMPLATES) {
        assert_eq!(
            message_payload(source, identifier, 10_143)?,
            message_payload(target, identifier, 10_143)?,
        );
        assert_eq!(
            object_header(source, identifier)?,
            object_header(target, identifier)?,
        );
    }
    assert_eq!(
        message_payload(source, FIRST_SECTION_IDENTIFIER, 777)?,
        message_payload(target, FIRST_SECTION_IDENTIFIER, 777)?,
    );
    assert_eq!(
        message_payload(source, FIRST_SECTION_IDENTIFIER, 778)?,
        message_payload(target, FIRST_SECTION_IDENTIFIER, 778)?,
    );
    assert_eq!(
        normalize_message_length(&object_header(source, FIRST_SECTION_IDENTIFIER)?, 1)?,
        normalize_message_length(&object_header(target, FIRST_SECTION_IDENTIFIER)?, 1)?,
    );
    Ok(())
}

#[test]
fn native_fixture_reads_all_fields_and_exact_transaction_contract() -> TestResult {
    let package = Package::open(fixture_path())?;
    let before = fixture_settings()?;
    assert_eq!(package.section_settings(SectionSelector::index(0))?, before);
    assert_eq!(
        package.section_settings(SectionSelector::name("Blank"))?,
        before
    );
    let source = package.source_bytes().to_vec();
    let source_pointer = package.source_bytes().as_ptr();

    let noop_edit = package.edit_section_settings(SectionSelector::index(0))?;
    assert_eq!(noop_edit.settings(), &before);
    let noop = noop_edit.set(before.clone())?.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);
    let replayed_noop = package.apply_section_settings(noop.patch())?;
    assert!(replayed_noop.patch().is_noop());
    assert_eq!(
        replayed_noop.package().source_bytes().as_ptr(),
        source_pointer
    );

    let after = changed_settings()?;
    let changed = package
        .edit_section_settings(SectionSelector::name("Blank"))?
        .set(after.clone())?
        .commit()?;
    assert_eq!(changed.patch().before(), &before);
    assert_eq!(changed.patch().after(), &after);
    assert_eq!(
        changed
            .package()
            .section_settings(SectionSelector::index(0))?,
        after
    );
    assert!(changed.diagnostics().changed());
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(changed.diagnostics().deleted_previews(), 0);
    assert!(changed.diagnostics().full_reparse_performed());
    assert_one_iwa_member_changed(&source, changed.package().source_bytes())?;
    assert_previews(changed.package().source_bytes(), true)?;

    let applied = package.apply_section_settings(changed.patch())?;
    assert_eq!(
        applied.package().source_bytes(),
        changed.package().source_bytes()
    );
    assert!(matches!(
        changed.package().apply_section_settings(changed.patch()),
        Err(Error::PatchConflict)
    ));
    let inverse = changed.patch().inverse();
    assert_eq!(inverse.before(), &after);
    assert_eq!(inverse.after(), &before);
    assert_eq!(inverse.inverse(), changed.patch().clone());
    assert!(matches!(
        package.apply_section_settings(&inverse),
        Err(Error::PatchConflict)
    ));
    let restored = changed.package().apply_section_settings(&inverse)?;
    assert_eq!(restored.package().source_bytes(), source);
    assert_eq!(
        restored
            .package()
            .section_settings(SectionSelector::index(0))?,
        before
    );
    Ok(())
}

#[test]
fn all_boolean_presence_states_preserve_name_and_pagination() -> TestResult {
    let states = [None, Some(false), Some(true)];
    for inherit in states {
        for first in states {
            for even_odd in states {
                for hide_first in states {
                    let expected = settings(
                        Some("Alpha"),
                        [inherit, first, even_odd, hide_first],
                        Some(Start::Unknown(7)),
                        Some(PageNumbering::Unknown(3)),
                        Some(42),
                    )?;
                    let second = settings(
                        Some("Beta"),
                        [None; 4],
                        Some(Start::RightPage),
                        Some(PageNumbering::ContinueFromPrevious),
                        Some(9),
                    )?;
                    let bytes = synthetic_package(&expected, &second)?;
                    let package = Package::from_bytes(&bytes)?;
                    assert_eq!(
                        package.section_settings(SectionSelector::index(0))?,
                        expected
                    );
                    let pointer = package.source_bytes().as_ptr();
                    let noop = package
                        .edit_section_settings(SectionSelector::name("Alpha"))?
                        .set(expected.clone())?
                        .commit()?;
                    assert!(noop.patch().is_noop());
                    assert_eq!(noop.package().source_bytes(), bytes);
                    assert_eq!(noop.package().source_bytes().as_ptr(), pointer);
                    assert_eq!(
                        noop.package().section_settings(SectionSelector::index(0))?,
                        expected
                    );
                    assert_eq!(
                        noop.package().section_settings(SectionSelector::index(1))?,
                        second
                    );
                }
            }
        }
    }
    Ok(())
}

#[test]
fn first_section_inheritance_changes_fail_before_template_traversal() -> TestResult {
    for (before_inherit, after_inherit) in [(false, true), (true, false)] {
        let before = settings(
            Some("First"),
            [Some(before_inherit), Some(false), Some(false), Some(false)],
            Some(Start::NextPage),
            Some(PageNumbering::ContinueFromPrevious),
            Some(1),
        )?;
        let second = fixture_settings()?;
        let bytes = synthetic_package(&before, &second)?;
        let package = Package::from_bytes(&bytes)?;
        let mut after = before.clone();
        after.set_inherit_previous_header_footer(Some(after_inherit));
        assert_eq!(
            package
                .edit_section_settings(SectionSelector::index(0))?
                .set(after)?
                .commit()
                .unwrap_err(),
            Error::UnsupportedDependency {
                path: Path::section(Position::new(0)),
                kind: DependencyKind::PreviousSectionTemplates,
            }
        );
        assert_eq!(package.source_bytes(), bytes);
    }
    Ok(())
}

#[test]
fn changed_aggregate_preserves_opaque_locality_and_inverts_exactly() -> TestResult {
    let before = settings(
        Some("Alpha"),
        [None, Some(false), Some(true), None],
        Some(Start::NextPage),
        Some(PageNumbering::ContinueFromPrevious),
        Some(1),
    )?;
    let untouched = settings(
        Some("Beta"),
        [Some(true), None, Some(false), Some(true)],
        Some(Start::RightPage),
        Some(PageNumbering::Restart),
        Some(9),
    )?;
    let after = settings(
        Some("A much longer aggregate section name"),
        [None, Some(true), None, Some(false)],
        Some(Start::Unknown(u32::MAX)),
        Some(PageNumbering::Unknown(u32::MAX)),
        Some(u32::MAX),
    )?;
    let bytes = synthetic_package(&before, &untouched)?;
    assert_previews(&bytes, true)?;
    let package = Package::from_bytes(&bytes)?;

    let commit = package
        .edit_section_settings(SectionSelector::name("Alpha"))?
        .set(after.clone())?
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.patch().path().position(), Some(Position::new(0)));
    assert_eq!(commit.patch().before(), &before);
    assert_eq!(commit.patch().after(), &after);
    assert_changed_locality(&bytes, commit.package().source_bytes())?;
    assert_eq!(
        commit
            .package()
            .section_settings(SectionSelector::index(1))?,
        untouched
    );

    let restored = commit
        .package()
        .apply_section_settings(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_previews(restored.package().source_bytes(), true)?;
    Ok(())
}

#[test]
fn selectors_and_legacy_sources_are_typed_and_failure_atomic() -> TestResult {
    let same = settings(Some("Same"), [None; 4], None, None, None)?;
    let bytes = synthetic_package(&same, &same)?;
    let package = Package::from_bytes(&bytes)?;

    assert!(matches!(
        package.edit_section_settings(SectionSelector::name("missing-private-selector")),
        Err(Error::NameNotFound)
    ));
    assert!(matches!(
        package.edit_section_settings(SectionSelector::index(2)),
        Err(Error::PositionNotFound { position }) if position == Position::new(2)
    ));
    assert!(matches!(
        package.edit_section_settings(SectionSelector::name("Same")),
        Err(Error::AmbiguousSelector { first, duplicate })
            if first == Position::new(0) && duplicate == Position::new(1)
    ));
    assert_eq!(package.source_bytes(), bytes);

    let first = settings(Some("Alpha"), [None; 4], None, None, None)?;
    let second = settings(Some("Beta"), [None; 4], None, None, None)?;
    let flat = synthetic_package(&first, &second)?;
    let legacy = legacy_package_bytes(&flat)?;
    let legacy_package = Package::from_bytes(&legacy)?;
    assert_eq!(
        legacy_package.section_settings(SectionSelector::index(0))?,
        first
    );
    let noop = legacy_package
        .edit_section_settings(SectionSelector::index(0))?
        .set(first.clone())?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().source_bytes(), legacy);

    let changed = settings(
        Some("Alpha"),
        [Some(true), None, None, None],
        None,
        None,
        None,
    )?;
    assert!(matches!(
        legacy_package
            .edit_section_settings(SectionSelector::index(0))?
            .set(changed)?
            .commit(),
        Err(Error::UnsupportedSource { .. })
    ));
    assert_eq!(legacy_package.source_bytes(), legacy);
    Ok(())
}

fn assert_malformed_fails_closed(payload: Vec<u8>, second: &Settings) -> TestResult {
    let bytes = synthetic_package_with_first_message(
        RawMessage {
            type_: SECTION_MESSAGE_TYPE,
            data: payload,
        },
        second,
    )?;
    let Ok(package) = Package::from_bytes(&bytes) else {
        return Ok(());
    };
    assert!(matches!(
        package.section_settings(SectionSelector::index(0)),
        Err(Error::InvalidSource { .. })
    ));
    assert!(matches!(
        package.edit_section_settings(SectionSelector::index(0)),
        Err(Error::InvalidSource { .. })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn malformed_selected_fields_fail_closed_before_publication() -> TestResult {
    let canonical = settings(
        Some("Alpha"),
        [Some(true), Some(false), Some(true), Some(false)],
        Some(Start::NextPage),
        Some(PageNumbering::Restart),
        Some(1),
    )?;
    let second = settings(Some("Beta"), [None; 4], None, None, None)?;
    let canonical_payload = section_payload(&canonical, FIRST_TEMPLATES, 7_777)?;

    for field in [17_u32, 18, 19, 20, 21, 22, 28] {
        let mut duplicate = canonical_payload.clone();
        append_varint_field(&mut duplicate, field, 0)?;
        assert_malformed_fails_closed(duplicate, &second)?;

        let mut wrong_wire = canonical_payload.clone();
        append_length_delimited_field(&mut wrong_wire, field, &[])?;
        assert_malformed_fails_closed(wrong_wire, &second)?;
    }
    let mut duplicate_name = canonical_payload.clone();
    append_length_delimited_field(&mut duplicate_name, 26, b"duplicate")?;
    assert_malformed_fails_closed(duplicate_name, &second)?;
    let mut wrong_name_wire = canonical_payload.clone();
    append_varint_field(&mut wrong_name_wire, 26, 1)?;
    assert_malformed_fails_closed(wrong_name_wire, &second)?;

    for field in [17_u32, 18, 19, 28] {
        let mut non_boolean = section_payload(
            &settings(None, [None; 4], None, None, None)?,
            FIRST_TEMPLATES,
            7_777,
        )?;
        append_varint_field(&mut non_boolean, field, 2)?;
        assert_malformed_fails_closed(non_boolean, &second)?;
    }

    let mut zero_page = section_payload(
        &settings(None, [None; 4], None, None, None)?,
        FIRST_TEMPLATES,
        7_777,
    )?;
    append_varint_field(&mut zero_page, 22, 0)?;
    assert_malformed_fails_closed(zero_page, &second)?;

    let mut noncanonical_varint = section_payload(
        &settings(None, [None; 4], None, None, None)?,
        FIRST_TEMPLATES,
        7_777,
    )?;
    encode_varint_into(&mut noncanonical_varint, u64::from(17_u32) << 3);
    noncanonical_varint.extend_from_slice(&[0x80, 0x00]);
    assert_malformed_fails_closed(noncanonical_varint, &second)?;

    let mut noncanonical_key = section_payload(
        &settings(None, [None; 4], None, None, None)?,
        FIRST_TEMPLATES,
        7_777,
    )?;
    noncanonical_key.extend_from_slice(&[0x88, 0x81, 0x00, 0x00]);
    assert_malformed_fails_closed(noncanonical_key, &second)?;

    let mut truncated = section_payload(
        &settings(None, [None; 4], None, None, None)?,
        FIRST_TEMPLATES,
        7_777,
    )?;
    encode_varint_into(&mut truncated, u64::from(28_u32) << 3);
    truncated.push(0x80);
    assert_malformed_fails_closed(truncated, &second)?;
    Ok(())
}

#[test]
fn public_values_are_send_sync_and_debug_redacts_private_content() -> TestResult {
    assert_type_send_sync::<Settings>();
    assert_type_send_sync::<Edit<'static>>();
    assert_type_send_sync::<Patch>();
    assert_type_send_sync::<Commit>();
    assert_type_send_sync::<Diagnostics>();
    assert_type_send_sync::<Error>();
    assert_type_send_sync::<LimitKind>();
    assert_type_send_sync::<Path>();
    assert_type_send_sync::<DependencyKind>();

    let before = settings(Some("Alpha"), [None; 4], None, None, None)?;
    let mut after = changed_settings()?;
    after.set_inherit_previous_header_footer(before.inherit_previous_header_footer());
    let private_name = after.name().unwrap_or_default().to_owned();
    let bytes = synthetic_package(&before, &fixture_settings()?)?;
    let package = Package::from_bytes(&bytes)?;
    let edit = package
        .edit_section_settings(SectionSelector::index(0))?
        .set(after)?;
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains(PRIVATE_MARKER));
    assert!(!edit_debug.contains(&private_name));
    assert!(!edit_debug.contains("Index/"));
    let commit = edit.commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    assert!(!patch_debug.contains(PRIVATE_MARKER));
    assert!(!patch_debug.contains(&private_name));
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains("fingerprint"));
    assert!(!patch_debug.contains("bytes"));
    Ok(())
}
