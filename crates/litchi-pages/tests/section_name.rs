use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use litchi_pages::{
    Limits, Package, Position, SectionNameError, SectionNameLimitKind, SectionSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const PRIVATE_NATIVE_MARKER: &str = "private-pages-native-marker-998244353";
const SECTION_MESSAGE_TYPE: u32 = 10_011;

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

fn section_message(name: &str, unknown_value: u64) -> TestResult<RawMessage> {
    let mut data = tp::SectionArchive {
        name: Some(name.to_owned()),
        ..tp::SectionArchive::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut data, 99, unknown_value)?;
    Ok(RawMessage {
        type_: SECTION_MESSAGE_TYPE,
        data,
    })
}

fn synthetic_package(first_name: &str, second_name: &str) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
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

    let first_section = ArchiveObject::new(
        FIRST_SECTION_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-section-payload".to_vec(),
            },
            section_message(first_name, 7_777)?,
            RawMessage {
                type_: 778,
                data: b"after-section-payload".to_vec(),
            },
        ],
    )?;
    let second_section = ArchiveObject::new(
        SECOND_SECTION_IDENTIFIER,
        vec![section_message(second_name, 8_888)?],
    )?;
    let document = component_with_unknown_header(
        vec![
            object(1, 10_000, root.encode_to_vec())?,
            object(BODY_IDENTIFIER, 2_001, body.encode_to_vec())?,
            first_section,
            second_section,
        ],
        FIRST_SECTION_IDENTIFIER,
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

fn fields_except_name(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse_with_limits(payload, WireLimits::default())?
        .fields()
        .filter(|field| field.number() != 26)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn section_snapshot(package: &Package) -> Vec<(Option<String>, String)> {
    package
        .sections()
        .iter()
        .map(|section| (section.name().map(str::to_owned), section.plain_text()))
        .collect()
}

fn expansive_name() -> String {
    (0..512)
        .filter_map(|offset| char::from_u32(0x4e00 + offset))
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
            // A preceding member changed length, so the central directory's
            // relative local-header offset is expected to move. All other
            // bytes, including extras, comments, attributes and raw name,
            // remain preservation authority.
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
fn native_fixture_renames_by_position_and_name_with_clear_and_empty() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let source_text = package.text()?;
    let source_sections = package.stats().section_count();

    let long_name = "Native Pages section name with a deliberately longer UTF-8 value";
    let mut by_position = package.edit_section_name(SectionSelector::index(0))?;
    assert_eq!(by_position.name(), Some("Blank"));
    by_position.set_name(Some(long_name))?;
    let longer = by_position.commit()?;
    assert_eq!(longer.package().sections()[0].name(), Some(long_name));
    assert_eq!(longer.package().text()?, source_text);
    assert_eq!(longer.package().stats().section_count(), source_sections);
    assert!(longer.diagnostics().changed());
    assert_eq!(longer.diagnostics().touched_components(), 1);
    assert!(longer.diagnostics().full_reparse_performed());
    assert_eq!(longer.patch().position().get(), 0);
    assert_eq!(longer.patch().before(), Some("Blank"));
    assert_eq!(longer.patch().after(), Some(long_name));

    let longer_package = longer.into_package();
    let mut by_name = longer_package.edit_section_name(SectionSelector::name(long_name))?;
    assert_eq!(by_name.name(), Some(long_name));
    by_name.set_name(Some("x"))?;
    let shorter = by_name.commit()?;
    assert_eq!(shorter.package().sections()[0].name(), Some("x"));

    let shorter_package = shorter.into_package();
    let mut clear = shorter_package.edit_section_name(SectionSelector::index(0))?;
    clear.clear_name();
    let cleared = clear.commit()?;
    assert_eq!(cleared.package().sections()[0].name(), None);

    let cleared_package = cleared.into_package();
    let mut empty_edit = cleared_package.edit_section_name(SectionSelector::index(0))?;
    empty_edit.set_name(Some(""))?;
    let empty_commit = empty_edit.commit()?;
    assert_eq!(empty_commit.package().sections()[0].name(), Some(""));
    assert_eq!(empty_commit.package().text()?, source_text);
    assert_eq!(
        empty_commit.package().stats().section_count(),
        source_sections
    );
    Ok(())
}

#[test]
fn selector_and_name_validation_fail_before_publication() -> TestResult<()> {
    let bytes = synthetic_package("Same", "Same")?;
    let package = Package::from_bytes(&bytes)?;

    assert!(matches!(
        package.edit_section_name(SectionSelector::name("Missing")),
        Err(SectionNameError::NameNotFound)
    ));
    assert!(matches!(
        package.edit_section_name(SectionSelector::index(2)),
        Err(SectionNameError::PositionNotFound { position })
            if position == Position::new(2)
    ));
    assert!(matches!(
        package.edit_section_name(SectionSelector::name("Same")),
        Err(SectionNameError::AmbiguousSelector { first, duplicate })
            if first == Position::new(0) && duplicate == Position::new(1)
    ));

    let mut edit = package.edit_section_name(SectionSelector::index(0))?;
    assert!(matches!(
        edit.set_name(Some("bad\0name")),
        Err(SectionNameError::InvalidName(_))
    ));
    assert_eq!(edit.name(), Some("Same"));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn no_op_reuses_exact_source_and_applies_to_legacy_source() -> TestResult<()> {
    let bytes = synthetic_package("Alpha", "Beta")?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();

    let mut edit = package.edit_section_name(SectionSelector::name("Alpha"))?;
    edit.set_name(Some("Alpha"))?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(
        noop.patch().source_fingerprint(),
        noop.patch().target_fingerprint()
    );
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(noop.package().source_bytes(), bytes);
    assert_eq!(noop.package().source_bytes().as_ptr(), source_pointer);
    assert_eq!(
        package
            .apply_section_name(noop.patch())?
            .package()
            .source_bytes(),
        bytes
    );

    let legacy = legacy_package_bytes(&bytes)?;
    let legacy_package = Package::from_bytes(&legacy)?;
    let mut legacy_noop_edit = legacy_package.edit_section_name(SectionSelector::index(0))?;
    legacy_noop_edit.set_name(Some("Alpha"))?;
    let legacy_noop_commit = legacy_noop_edit.commit()?;
    assert!(legacy_noop_commit.patch().is_noop());
    assert_eq!(legacy_noop_commit.package().source_bytes(), legacy);
    assert_eq!(
        legacy_package
            .apply_section_name(legacy_noop_commit.patch())?
            .package()
            .source_bytes(),
        legacy
    );

    let mut legacy_changed = legacy_package.edit_section_name(SectionSelector::index(0))?;
    legacy_changed.set_name(Some("Changed"))?;
    assert!(matches!(
        legacy_changed.commit(),
        Err(SectionNameError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn changed_name_preserves_unrelated_content_and_reversible_patch() -> TestResult<()> {
    let bytes = synthetic_package("Alpha", "Beta")?;
    let package = Package::from_bytes(&bytes)?;
    let source_sections = section_snapshot(&package);
    let source_text = package.text()?;
    let source_first_payload =
        message_payload(&bytes, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let source_first_header = object_header(&bytes, FIRST_SECTION_IDENTIFIER)?;
    let source_second_payload =
        message_payload(&bytes, SECOND_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let source_body_payload = message_payload(&bytes, BODY_IDENTIFIER, 2_001)?;
    let replacement = expansive_name();

    let mut edit = package.edit_section_name(SectionSelector::index(0))?;
    edit.set_name(Some(&replacement))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().position().get(), 0);
    assert_eq!(commit.patch().before(), Some("Alpha"));
    assert_eq!(commit.patch().after(), Some(replacement.as_str()));
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(commit.package().text()?, source_text);
    assert_eq!(
        commit.package().sections()[0].name(),
        Some(replacement.as_str())
    );
    assert_eq!(commit.package().sections()[1].name(), Some("Beta"));
    assert_eq!(
        commit.package().sections()[0].plain_text(),
        source_sections[0].1
    );
    assert_eq!(
        commit.package().sections()[1].plain_text(),
        source_sections[1].1
    );

    let target = commit.package().source_bytes();
    let target_first_payload =
        message_payload(target, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let target_first_header = object_header(target, FIRST_SECTION_IDENTIFIER)?;
    assert_eq!(
        fields_except_name(&target_first_payload)?,
        fields_except_name(&source_first_payload)?
    );
    assert!(
        WireView::parse(&target_first_payload)?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        normalize_message_length(&target_first_header, 1)?,
        normalize_message_length(&source_first_header, 1)?,
        "only the renamed message length may change in the opaque archive header"
    );
    assert!(
        WireView::parse(&target_first_header)?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        message_payload(target, SECOND_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?,
        source_second_payload
    );
    assert_eq!(
        message_payload(target, BODY_IDENTIFIER, 2_001)?,
        source_body_payload
    );
    assert_untouched_zip_members(&bytes, target)?;

    let applied = package.apply_section_name(commit.patch())?;
    assert_eq!(applied.package().source_bytes(), target);

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), Some(replacement.as_str()));
    assert_eq!(inverse.after(), Some("Alpha"));
    assert_eq!(
        inverse.source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(
        inverse.target_fingerprint(),
        commit.patch().source_fingerprint()
    );
    let restored = commit.package().apply_section_name(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(section_snapshot(restored.package()), source_sections);

    let unrelated = Package::from_bytes(&synthetic_package("Other", "Beta")?)?;
    assert!(matches!(
        unrelated.apply_section_name(commit.patch()),
        Err(SectionNameError::PatchConflict)
    ));

    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains(PRIVATE_NATIVE_MARKER));
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("fingerprint"));
    assert!(!debug.contains("bytes"));

    assert_send_sync(&package);
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(&commit.diagnostics());
    assert_type_send_sync::<SectionNameError>();
    Ok(())
}

#[test]
fn changed_name_respects_retained_tight_output_limit() -> TestResult<()> {
    let bytes = synthetic_package("Alpha", "Beta")?;
    let input_bytes = u64::try_from(bytes.len())?;
    let limits = Limits::new(input_bytes, 8, 1_024 * 1_024, 1_024 * 1_024, 1_024 * 1_024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let replacement = expansive_name();
    let mut edit = package.edit_section_name(SectionSelector::index(0))?;
    edit.set_name(Some(&replacement))?;
    assert!(matches!(
        edit.commit(),
        Err(SectionNameError::LimitExceeded {
            kind: SectionNameLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}
