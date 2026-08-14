use std::io;
use std::sync::{Arc, Barrier};

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsp, tswp};
use litchi_pages::{
    Limits, Package, PackageError, Position, SectionSelector, SectionTextError,
    SectionTextLimitKind, TextPosition, TextSpan,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const THIRD_SECTION_IDENTIFIER: u64 = 45;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const UNRELATED_STORAGE_MESSAGE_TYPE: u32 = 2_002;
const SECTION_MESSAGE_TYPE: u32 = 10_011;
const PRIVATE_NATIVE_MARKER: &str = "private-pages-body-marker-2147483647";
const UNRELATED_STORAGE_MARKER: &[u8] = b"unrelated-type-2002-storage-sibling";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn assert_send_sync<T: Send + Sync>(_: &T) {}
fn assert_type_send_sync<T: Send + Sync>() {}

fn fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/basic.pages")
}

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

fn section_message(name: &str, unknown: u64) -> TestResult<RawMessage> {
    let mut data = tp::SectionArchive {
        name: Some(name.to_owned()),
        ..tp::SectionArchive::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut data, 99, unknown)?;
    Ok(RawMessage {
        type_: SECTION_MESSAGE_TYPE,
        data,
    })
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_text(["First😀\u{0004}Se", "cond東京\u{0004}Third"])
}

fn synthetic_package_with_text<const N: usize>(fragments: [&str; N]) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        section: Some(reference(FIRST_SECTION_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let text = fragments.into_iter().map(str::to_owned).collect::<Vec<_>>();
    let mut boundaries = vec![0_u32];
    let mut utf16_offset = 0usize;
    for character in text.iter().flat_map(|fragment| fragment.chars()) {
        utf16_offset = utf16_offset
            .checked_add(character.len_utf16())
            .ok_or_else(|| io::Error::other("synthetic UTF-16 length overflow"))?;
        if character == '\u{0004}' {
            boundaries.push(u32::try_from(utf16_offset)?);
        }
    }
    if boundaries.len() != 3 {
        return Err(io::Error::other("synthetic body requires exactly three sections").into());
    }
    let encoded_body = tswp::StorageArchive {
        text,
        table_section: Some(tswp::ObjectAttributeTable {
            entries: vec![
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(reference(FIRST_SECTION_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: boundaries[1],
                    object: Some(reference(SECOND_SECTION_IDENTIFIER)),
                },
                tswp::object_attribute_table::ObjectAttribute {
                    character_index: boundaries[2],
                    object: Some(reference(THIRD_SECTION_IDENTIFIER)),
                },
            ],
        }),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    // Interleave an unknown field between the two repeated text records so a
    // changed splice must preserve both the record ordering and raw unknown
    // bytes rather than normalize the complete message.
    let encoded_view = WireView::parse(&encoded_body)?;
    let mut body_payload = Vec::with_capacity(encoded_body.len().saturating_add(8));
    let mut text_fields = 0usize;
    for field in encoded_view.fields() {
        body_payload.extend_from_slice(field.raw());
        if field.number() == 3 {
            text_fields += 1;
            if text_fields == 1 {
                litchi_iwa_common::wire::append_varint_field(&mut body_payload, 99, 123_456)?;
            }
        }
    }

    let body = ArchiveObject::new(
        BODY_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-body-message".to_vec(),
            },
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: body_payload,
            },
            RawMessage {
                type_: UNRELATED_STORAGE_MESSAGE_TYPE,
                data: UNRELATED_STORAGE_MARKER.to_vec(),
            },
            RawMessage {
                type_: 778,
                data: b"after-body-message".to_vec(),
            },
        ],
    )?;
    let first = ArchiveObject::new(
        FIRST_SECTION_IDENTIFIER,
        vec![section_message("First", 11)?],
    )?;
    let second = ArchiveObject::new(
        SECOND_SECTION_IDENTIFIER,
        vec![section_message("Middle", 22)?],
    )?;
    let third = ArchiveObject::new(THIRD_SECTION_IDENTIFIER, vec![section_message("Last", 33)?])?;
    let document = component_with_unknown_header(
        vec![
            object(1, 10_000, root.encode_to_vec())?,
            body,
            first,
            second,
            third,
        ],
        BODY_IDENTIFIER,
    )?;
    let unrelated = component(vec![object(
        99,
        777,
        PRIVATE_NATIVE_MARKER.as_bytes().to_vec(),
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document.as_slice()),
            (UNRELATED_MEMBER, unrelated.as_slice()),
        ],
        Limits::default(),
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

fn package_with_hidden_bookmark_reference() -> TestResult<Vec<u8>> {
    let bytes = synthetic_package()?;
    let stream = document_stream(&bytes)?;
    let mut archive = Archive::parse(&stream)?;
    let body = archive
        .object_mut(BODY_IDENTIFIER)
        .ok_or_else(|| io::Error::other("missing synthetic body"))?;
    let message_index = body
        .messages
        .iter()
        .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic body message"))?;
    let mut storage = tswp::StorageArchive::decode(body.messages[message_index].data.as_slice())?;
    storage.table_bookmark = Some(tswp::ObjectAttributeTable {
        entries: vec![
            tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: None,
            },
            tswp::object_attribute_table::ObjectAttribute {
                character_index: 2,
                object: Some(reference(90)),
            },
        ],
    });
    body.replace_message(
        message_index,
        RawMessage {
            type_: STORAGE_MESSAGE_TYPE,
            data: storage.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(&bytes)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn package_with_duplicate_known_table() -> TestResult<Vec<u8>> {
    let bytes = synthetic_package()?;
    let stream = document_stream(&bytes)?;
    let mut archive = Archive::parse(&stream)?;
    let body = archive
        .object_mut(BODY_IDENTIFIER)
        .ok_or_else(|| io::Error::other("missing synthetic body"))?;
    let message_index = body
        .messages
        .iter()
        .position(|message| message.type_ == STORAGE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic body message"))?;
    let mut payload = body.messages[message_index].data.clone();
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 11, &[])?;
    litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 11, &[])?;
    body.replace_message(
        message_index,
        RawMessage {
            type_: STORAGE_MESSAGE_TYPE,
            data: payload,
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(&bytes)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic Pages document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
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
    let (header_length, prefix_length) = decode_varint_from_bytes(&stream[header_offset..])?;
    let start = header_offset + prefix_length;
    let end = start + usize::try_from(header_length)?;
    assert_eq!(end, data_offset);
    Ok(stream[start..end].to_vec())
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

fn body_fields_except_text_and_sections(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| !matches!(field.number(), 3 | 17))
        .map(|field| field.raw().to_vec())
        .collect())
}

fn text_field_raw(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == 3)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn field_numbers(payload: &[u8]) -> TestResult<Vec<u32>> {
    Ok(WireView::parse(payload)?
        .fields()
        .map(litchi_iwa_common::wire::WireFieldView::number)
        .collect())
}

fn section_snapshot(package: &Package) -> TestResult<Vec<(Option<String>, String)>> {
    package
        .sections()
        .iter()
        .enumerate()
        .map(|(index, section)| {
            Ok((
                section.name().map(str::to_owned),
                package
                    .section_text(SectionSelector::index(index))?
                    .to_owned(),
            ))
        })
        .collect()
}

fn assert_untouched_zip_members(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let before_entries = before_catalog.iter().collect::<Vec<_>>();
    let after_entries = after_catalog.iter().collect::<Vec<_>>();
    assert_eq!(before_entries.len(), after_entries.len());
    let mut changed = 0usize;
    for (before_entry, after_entry) in before_entries.into_iter().zip(after_entries) {
        assert_eq!(before_entry.name(), after_entry.name());
        assert_eq!(before_entry.raw_name(), after_entry.raw_name());
        if before_entry.data() == after_entry.data() {
            assert_eq!(before_entry.metadata(), after_entry.metadata());
            assert_eq!(
                before_entry.raw_record().local_record(),
                after_entry.raw_record().local_record()
            );
            let mut before_central = before_entry
                .raw_record()
                .central_directory_record()
                .to_vec();
            let mut after_central = after_entry.raw_record().central_directory_record().to_vec();
            before_central[42..46].fill(0);
            after_central[42..46].fill(0);
            assert_eq!(before_central, after_central);
        } else {
            changed += 1;
            assert_eq!(before_entry.name(), DOCUMENT_MEMBER);
        }
    }
    assert_eq!(changed, 1);
    Ok(())
}

#[test]
fn selector_first_astral_edit_shifts_only_following_boundaries() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.section_text(SectionSelector::name("First"))?,
        "First😀"
    );
    assert_eq!(
        package.section_text(SectionSelector::index(1))?,
        "Second東京"
    );
    assert_eq!(
        package.section_text(SectionSelector::name("Last"))?,
        "Third"
    );
    assert!(matches!(
        package.edit_body_text(),
        Err(SectionTextError::BodySectionCount { actual: 3 })
    ));

    let span = TextSpan::from_utf16_indexes(1, 1)?;
    let mut edit = package.edit_section_text(SectionSelector::name("Middle"))?;
    assert_eq!(edit.position(), Position::new(1));
    assert_eq!(edit.text(), "Second東京");
    edit.insert(TextPosition::from_utf16_code_units(1), "😀")?;
    assert_eq!(edit.span(), Some(span));
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.patch().span(), span);
    assert_eq!(commit.patch().before(), "Second東京");
    assert_eq!(commit.patch().after(), "S😀econd東京");
    assert_eq!(
        section_snapshot(commit.package())?,
        vec![
            (Some("First".to_owned()), "First😀".to_owned()),
            (Some("Middle".to_owned()), "S😀econd東京".to_owned()),
            (Some("Last".to_owned()), "Third".to_owned()),
        ]
    );

    let body = tswp::StorageArchive::decode(
        message_payload(
            commit.package().source_bytes(),
            BODY_IDENTIFIER,
            STORAGE_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    let indexes = body
        .table_section
        .ok_or_else(|| io::Error::other("missing section table"))?
        .entries
        .into_iter()
        .map(|entry| entry.character_index)
        .collect::<Vec<_>>();
    assert_eq!(indexes, [0, 8, 19]);
    Ok(())
}

#[test]
fn no_op_shares_exact_and_legacy_sources() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let mut edit = package.edit_section_text(SectionSelector::index(0))?;
    edit.set("First😀")?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);
    assert_eq!(
        package
            .apply_section_text(noop.patch())?
            .package()
            .source_bytes()
            .as_ptr(),
        source_pointer
    );

    let legacy = legacy_package_bytes(&bytes)?;
    let legacy_package = Package::from_bytes(&legacy)?;
    let mut legacy_noop_edit = legacy_package.edit_section_text(SectionSelector::index(2))?;
    legacy_noop_edit.set("Third")?;
    let legacy_noop_commit = legacy_noop_edit.commit()?;
    assert!(legacy_noop_commit.patch().is_noop());
    assert_eq!(legacy_noop_commit.package().source_bytes(), legacy);

    let mut changed = legacy_package.edit_section_text(SectionSelector::index(2))?;
    changed.set("Changed")?;
    assert!(matches!(
        changed.commit(),
        Err(SectionTextError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn single_section_body_convenience_resolves_without_native_identity() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    assert_eq!(package.sections().len(), 1);
    let original = package.section_text(SectionSelector::index(0))?;
    let mut edit = package.edit_body_text()?;
    assert_eq!(edit.position(), Position::new(0));
    assert_eq!(edit.text(), original);
    edit.set(original)?;
    let commit = edit.commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.package().source_bytes().as_ptr(),
        package.source_bytes().as_ptr()
    );
    Ok(())
}

#[test]
fn changed_text_preserves_unknown_data_headers_and_zip_and_is_reversible() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    let source_snapshot = section_snapshot(&package)?;
    let source_body = message_payload(&bytes, BODY_IDENTIFIER, STORAGE_MESSAGE_TYPE)?;
    let source_unrelated_storage =
        message_payload(&bytes, BODY_IDENTIFIER, UNRELATED_STORAGE_MESSAGE_TYPE)?;
    let source_before_message = message_payload(&bytes, BODY_IDENTIFIER, 777)?;
    let source_after_message = message_payload(&bytes, BODY_IDENTIFIER, 778)?;
    let source_header = object_header(&bytes, BODY_IDENTIFIER)?;
    let source_first_section =
        message_payload(&bytes, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let source_text_fields = text_field_raw(&source_body)?;

    let mut edit = package.edit_section_text(SectionSelector::index(1))?;
    edit.replace(TextSpan::from_utf16_indexes(0, 1)?, "astral😀replacement")?;
    let edit_debug = format!("{edit:?}");
    assert!(!edit_debug.contains("Second"));
    assert!(!edit_debug.contains("replacement"));
    assert!(!edit_debug.contains("Index/"));
    let commit = edit.commit()?;
    let target = commit.package().source_bytes();
    let target_body = message_payload(target, BODY_IDENTIFIER, STORAGE_MESSAGE_TYPE)?;
    let target_header = object_header(target, BODY_IDENTIFIER)?;
    assert_eq!(
        body_fields_except_text_and_sections(&target_body)?,
        body_fields_except_text_and_sections(&source_body)?
    );
    assert_eq!(field_numbers(&target_body)?, field_numbers(&source_body)?);
    assert!(
        WireView::parse(&target_body)?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        normalize_message_length(&target_header, 1)?,
        normalize_message_length(&source_header, 1)?
    );
    assert!(
        WireView::parse(&target_header)?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        message_payload(target, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?,
        source_first_section
    );
    assert_eq!(
        message_payload(target, BODY_IDENTIFIER, 777)?,
        source_before_message
    );
    assert_eq!(
        message_payload(target, BODY_IDENTIFIER, 778)?,
        source_after_message
    );
    assert_eq!(
        message_payload(target, BODY_IDENTIFIER, UNRELATED_STORAGE_MESSAGE_TYPE)?,
        source_unrelated_storage
    );
    let target_text_fields = text_field_raw(&target_body)?;
    assert_eq!(target_text_fields.len(), source_text_fields.len());
    assert_eq!(target_text_fields[1], source_text_fields[1]);
    assert_untouched_zip_members(&bytes, target)?;

    let applied = package.apply_section_text(commit.patch())?;
    assert_eq!(applied.package().source_bytes(), target);
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), commit.patch().after());
    assert_eq!(inverse.after(), commit.patch().before());
    let restored = commit.package().apply_section_text(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(section_snapshot(restored.package())?, source_snapshot);

    let equivalent_bytes = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            "Data/sentinel.bin",
            b"different unrelated ZIP sentinel",
        )],
        Limits::default(),
    )?;
    let equivalent = Package::from_bytes(&equivalent_bytes)?;
    assert_eq!(section_snapshot(&equivalent)?, source_snapshot);
    assert!(matches!(
        equivalent.apply_section_text(commit.patch()),
        Err(SectionTextError::PatchConflict)
    ));

    let unrelated = Package::from_bytes(&synthetic_package()?)?;
    let mut unrelated_edit = unrelated.edit_section_text(SectionSelector::index(0))?;
    unrelated_edit.insert(TextPosition::ZERO, "different")?;
    let unrelated_commit = unrelated_edit.commit()?;
    assert!(matches!(
        unrelated_commit
            .package()
            .apply_section_text(commit.patch()),
        Err(SectionTextError::PatchConflict)
    ));

    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains(PRIVATE_NATIVE_MARKER));
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("Second"));
    assert!(!debug.contains("replacement"));
    let commit_debug = format!("{commit:?}");
    assert!(!commit_debug.contains("Second"));
    assert!(!commit_debug.contains("replacement"));
    assert!(!commit_debug.contains("Index/"));
    assert_send_sync(&package);
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(commit.diagnostics());
    assert_type_send_sync::<SectionTextError>();
    Ok(())
}

#[test]
fn concurrent_section_text_commits_are_isolated_and_source_remains_immutable() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Arc::new(Package::from_bytes(&bytes)?);
    let source_pointer = package.source_bytes().as_ptr();
    let barrier = Arc::new(Barrier::new(3));

    let first_package = Arc::clone(&package);
    let first_barrier = Arc::clone(&barrier);
    let first = std::thread::spawn(move || {
        first_barrier.wait();
        let mut edit = first_package.edit_section_text(SectionSelector::index(0))?;
        edit.set("parallel first")?;
        edit.commit()
    });

    let last_package = Arc::clone(&package);
    let last_barrier = Arc::clone(&barrier);
    let last = std::thread::spawn(move || {
        last_barrier.wait();
        let mut edit = last_package.edit_section_text(SectionSelector::index(2))?;
        edit.set("parallel last")?;
        edit.commit()
    });

    barrier.wait();
    let first = first
        .join()
        .map_err(|_| io::Error::other("first section editor panicked"))??;
    let last = last
        .join()
        .map_err(|_| io::Error::other("last section editor panicked"))??;

    assert_eq!(
        first.package().section_text(SectionSelector::index(0))?,
        "parallel first"
    );
    assert_eq!(
        first.package().section_text(SectionSelector::index(1))?,
        "Second東京"
    );
    assert_eq!(
        first.package().section_text(SectionSelector::index(2))?,
        "Third"
    );
    assert_eq!(
        last.package().section_text(SectionSelector::index(0))?,
        "First😀"
    );
    assert_eq!(
        last.package().section_text(SectionSelector::index(1))?,
        "Second東京"
    );
    assert_eq!(
        last.package().section_text(SectionSelector::index(2))?,
        "parallel last"
    );
    assert_ne!(
        first.package().source_bytes(),
        last.package().source_bytes()
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);

    let first_restored = first
        .package()
        .apply_section_text(&first.patch().inverse())?;
    assert_eq!(first_restored.package().source_bytes(), bytes);
    let last_restored = last.package().apply_section_text(&last.patch().inverse())?;
    assert_eq!(last_restored.package().source_bytes(), bytes);
    assert!(matches!(
        first.package().apply_section_text(last.patch()),
        Err(SectionTextError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn selectors_spans_reserved_markers_and_dependent_content_are_typed() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    assert!(matches!(
        package.edit_section_text(SectionSelector::name("Missing")),
        Err(SectionTextError::NameNotFound)
    ));
    assert!(matches!(
        package.edit_section_text(SectionSelector::index(3)),
        Err(SectionTextError::PositionNotFound { position }) if position == Position::new(3)
    ));

    let mut edit = package.edit_section_text(SectionSelector::index(0))?;
    assert!(matches!(
        edit.insert(TextPosition::ZERO, "bad\u{0004}break"),
        Err(SectionTextError::SectionBreakReplacement)
    ));
    assert!(matches!(
        edit.insert(TextPosition::ZERO, "bad\u{000e}footnote"),
        Err(SectionTextError::FootnoteAnchorReplacement)
    ));
    assert!(matches!(
        edit.insert(TextPosition::ZERO, "bad\u{fffc}object"),
        Err(SectionTextError::ObjectMarkerReplacement)
    ));
    assert!(matches!(
        edit.delete(TextSpan::from_utf16_indexes(100, 100)?),
        Err(SectionTextError::SpanOutOfBounds { .. })
    ));
    assert!(matches!(
        edit.delete(TextSpan::from_utf16_indexes(6, 6)?),
        Err(SectionTextError::SurrogateBoundary { position })
            if position == TextPosition::from_utf16_code_units(6)
    ));
    edit.insert(TextPosition::ZERO, "valid")?;
    assert!(matches!(
        edit.insert(TextPosition::ZERO, "second"),
        Err(SectionTextError::OperationAlreadyStaged)
    ));

    let dependent_bytes =
        synthetic_package_with_text(["First\u{000e}\u{fffc}\u{0004}Se", "cond東京\u{0004}Third"])?;
    let dependent_package = Package::from_bytes(&dependent_bytes)?;
    let mut dependent_edit = dependent_package.edit_section_text(SectionSelector::index(0))?;
    assert!(matches!(
        dependent_edit.delete(TextSpan::from_utf16_indexes(5, 6)?),
        Err(SectionTextError::DependentContent)
    ));
    let mut object_edit = dependent_package.edit_section_text(SectionSelector::index(0))?;
    assert!(matches!(
        object_edit.delete(TextSpan::from_utf16_indexes(6, 7)?),
        Err(SectionTextError::DependentContent)
    ));

    let referenced_bytes = package_with_hidden_bookmark_reference()?;
    let referenced_package = Package::from_bytes(&referenced_bytes)?;
    let referenced_source = referenced_package.source_bytes().to_vec();
    let mut referenced_edit = referenced_package.edit_section_text(SectionSelector::index(0))?;
    referenced_edit.delete(TextSpan::from_utf16_indexes(2, 3)?)?;
    assert!(matches!(
        referenced_edit.commit(),
        Err(SectionTextError::DependentContent)
    ));
    assert_eq!(referenced_package.source_bytes(), referenced_source);

    let duplicate_name_bytes = {
        let duplicate_source = synthetic_package()?;
        let stream = document_stream(&duplicate_source)?;
        let mut archive = Archive::parse(&stream)?;
        let second = archive
            .object_mut(SECOND_SECTION_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing second section"))?;
        second.replace_message(
            0,
            RawMessage {
                type_: SECTION_MESSAGE_TYPE,
                data: tp::SectionArchive {
                    name: Some("First".to_owned()),
                    ..tp::SectionArchive::default()
                }
                .encode_to_vec(),
            },
        )?;
        let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
        let catalog = Catalog::from_bytes(&duplicate_source)?;
        catalog.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                DOCUMENT_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?
    };
    let duplicate_name_package = Package::from_bytes(&duplicate_name_bytes)?;
    assert!(matches!(
        duplicate_name_package.edit_section_text(SectionSelector::name("First")),
        Err(SectionTextError::AmbiguousSelector { first, duplicate })
            if first == Position::new(0) && duplicate == Position::new(1)
    ));
    Ok(())
}

#[test]
fn set_clear_and_delete_preserve_neighboring_sections() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;

    let mut set_edit = package.edit_section_text(SectionSelector::index(2))?;
    set_edit.set("尾😀")?;
    let set_commit = set_edit.commit()?;
    assert_eq!(
        set_commit
            .package()
            .section_text(SectionSelector::index(2))?,
        "尾😀"
    );
    assert_eq!(
        set_commit
            .package()
            .section_text(SectionSelector::index(0))?,
        "First😀"
    );
    assert_eq!(
        set_commit
            .package()
            .section_text(SectionSelector::index(1))?,
        "Second東京"
    );

    let set_package = set_commit.into_package();
    let mut clear_edit = set_package.edit_section_text(SectionSelector::index(0))?;
    clear_edit.clear()?;
    let clear_commit = clear_edit.commit()?;
    assert_eq!(
        clear_commit
            .package()
            .section_text(SectionSelector::index(0))?,
        ""
    );
    assert_eq!(
        clear_commit
            .package()
            .section_text(SectionSelector::index(1))?,
        "Second東京"
    );

    let clear_package = clear_commit.into_package();
    let mut delete_edit = clear_package.edit_section_text(SectionSelector::index(1))?;
    delete_edit.delete(TextSpan::from_utf16_indexes(6, 8)?)?;
    let delete_commit = delete_edit.commit()?;
    assert_eq!(
        delete_commit
            .package()
            .section_text(SectionSelector::index(1))?,
        "Second"
    );
    assert_eq!(
        delete_commit
            .package()
            .section_text(SectionSelector::index(2))?,
        "尾😀"
    );
    Ok(())
}

#[test]
fn changed_text_respects_retained_output_limit() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let input_bytes = u64::try_from(bytes.len())?;
    let limits = Limits::new(input_bytes, 8, 1_024 * 1_024, 1_024 * 1_024, 1_024 * 1_024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut edit = package.edit_section_text(SectionSelector::index(1))?;
    edit.insert(TextPosition::ZERO, &"expanded".repeat(1_024))?;
    assert!(matches!(
        edit.commit(),
        Err(SectionTextError::LimitExceeded {
            kind: SectionTextLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn malformed_known_text_tables_fail_closed_before_lazy_projection() -> TestResult<()> {
    let bytes = package_with_duplicate_known_table()?;
    let error = Package::from_bytes(&bytes)
        .err()
        .ok_or_else(|| io::Error::other("duplicate known tables must fail bounded ingress"))?;
    assert!(matches!(&error, PackageError::InvalidFormat(_)));
    assert!(error.to_string().contains("failed bounded validation"));
    Ok(())
}
