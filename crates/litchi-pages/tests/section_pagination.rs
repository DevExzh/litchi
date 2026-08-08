use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{WireLimits, decode_varint_from_bytes, encode_varint_into, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsp, tswp};
use litchi_pages::{
    Limits, Package, Position, SectionPaginationError, SectionPaginationLimitKind, SectionSelector,
    section::{PageNumber, PageNumbering, Pagination, Start},
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_SECTION_IDENTIFIER: u64 = 43;
const SECOND_SECTION_IDENTIFIER: u64 = 44;
const PRIVATE_NATIVE_MARKER: &str = "private-pages-pagination-marker-998244353";
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

fn pagination(
    start: Option<Start>,
    numbering: Option<PageNumbering>,
    page: Option<u32>,
) -> TestResult<Pagination> {
    let mut value = Pagination::new();
    value.set_start(start)?;
    value.set_page_numbering(numbering)?;
    value.set_starting_page_number(page.map(PageNumber::new).transpose()?);
    Ok(value)
}

fn section_message(
    name: &str,
    pagination: Pagination,
    unknown_value: u64,
) -> TestResult<RawMessage> {
    let mut data = tp::SectionArchive {
        section_start_kind: pagination.start().map(Start::as_raw),
        section_page_number_kind: pagination.page_numbering().map(PageNumbering::as_raw),
        section_page_number_start: pagination.starting_page_number().map(PageNumber::get),
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

fn synthetic_package(
    first_name: &str,
    first_pagination: Pagination,
    second_name: &str,
    second_pagination: Pagination,
) -> TestResult<Vec<u8>> {
    synthetic_package_with_first_message(
        first_name,
        section_message(first_name, first_pagination, 7_777)?,
        second_name,
        second_pagination,
    )
}

fn synthetic_package_with_first_message(
    _first_name: &str,
    first_message: RawMessage,
    second_name: &str,
    second_pagination: Pagination,
) -> TestResult<Vec<u8>> {
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
            first_message,
            RawMessage {
                type_: 778,
                data: b"after-section-payload".to_vec(),
            },
        ],
    )?;
    let second_section = ArchiveObject::new(
        SECOND_SECTION_IDENTIFIER,
        vec![section_message(second_name, second_pagination, 8_888)?],
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

fn fields_except_pagination(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse_with_limits(payload, WireLimits::default())?
        .fields()
        .filter(|field| !matches!(field.number(), 20..=22))
        .map(|field| field.raw().to_vec())
        .collect())
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
fn native_fixture_reads_and_edits_pagination_by_semantic_selector() -> TestResult<()> {
    let package = Package::open(fixture_path())?;
    let original = pagination(
        Some(Start::NextPage),
        Some(PageNumbering::ContinueFromPrevious),
        Some(1),
    )?;
    assert_eq!(
        package.section_pagination(SectionSelector::index(0))?,
        original
    );
    let source_text = package.text()?;
    let source_name = package.sections()[0].name().map(str::to_owned);

    let mut edit = package.edit_section_pagination(SectionSelector::name("Blank"))?;
    edit.set_start(Some(Start::RightPage))?;
    edit.set_page_numbering(Some(PageNumbering::Restart))?;
    edit.set_starting_page_number(Some(PageNumber::new(7)?));
    let commit = edit.commit()?;
    let expected = pagination(
        Some(Start::RightPage),
        Some(PageNumbering::Restart),
        Some(7),
    )?;
    assert_eq!(
        commit
            .package()
            .section_pagination(SectionSelector::index(0))?,
        expected
    );
    assert_eq!(commit.package().text()?, source_text);
    assert_eq!(
        commit.package().sections()[0].name().map(str::to_owned),
        source_name
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    Ok(())
}

#[test]
fn all_presence_states_round_trip_and_noops_share_source() -> TestResult<()> {
    let starts = [None, Some(Start::NextPage), Some(Start::LeftPage)];
    let numberings = [
        None,
        Some(PageNumbering::ContinueFromPrevious),
        Some(PageNumbering::Restart),
    ];
    let pages = [None, Some(1), Some(42)];

    for start in starts {
        for numbering in numberings {
            for page in pages {
                let expected = pagination(start, numbering, page)?;
                let bytes = synthetic_package("Alpha", expected, "Beta", Pagination::new())?;
                let package = Package::from_bytes(&bytes)?;
                assert_eq!(
                    package.section_pagination(SectionSelector::index(0))?,
                    expected
                );
                let source_pointer = package.source_bytes().as_ptr();
                let mut edit = package.edit_section_pagination(SectionSelector::index(0))?;
                edit.set_pagination(expected)?;
                let commit = edit.commit()?;
                assert!(commit.patch().is_noop());
                assert_eq!(commit.package().source_bytes(), bytes);
                assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);
                assert!(!commit.diagnostics().changed());
                assert_eq!(commit.diagnostics().touched_components(), 0);
                assert!(!commit.diagnostics().full_reparse_performed());
            }
        }
    }
    Ok(())
}

#[test]
fn selector_validation_and_malformed_wire_fail_before_publication() -> TestResult<()> {
    let empty = Pagination::new();
    let selector_bytes = synthetic_package("Same", empty, "Same", empty)?;
    let selector_package = Package::from_bytes(&selector_bytes)?;
    assert!(matches!(
        selector_package.edit_section_pagination(SectionSelector::name("Missing")),
        Err(SectionPaginationError::NameNotFound)
    ));
    assert!(matches!(
        selector_package.edit_section_pagination(SectionSelector::index(2)),
        Err(SectionPaginationError::PositionNotFound { position })
            if position == Position::new(2)
    ));
    assert!(matches!(
        selector_package.edit_section_pagination(SectionSelector::name("Same")),
        Err(SectionPaginationError::AmbiguousSelector { first, duplicate })
            if first == Position::new(0) && duplicate == Position::new(1)
    ));
    let mut edit = selector_package.edit_section_pagination(SectionSelector::index(0))?;
    assert!(matches!(
        edit.set_start(Some(Start::Unknown(0))),
        Err(SectionPaginationError::InvalidPagination(_))
    ));
    assert_eq!(edit.pagination(), empty);
    assert_eq!(selector_package.source_bytes(), selector_bytes);

    let canonical = section_message("Alpha", empty, 7_777)?;
    let mut malformed_payloads = Vec::new();

    let mut duplicate = canonical.data.clone();
    litchi_iwa_common::wire::append_varint_field(&mut duplicate, 20, 1)?;
    litchi_iwa_common::wire::append_varint_field(&mut duplicate, 20, 2)?;
    malformed_payloads.push(duplicate);

    let mut wrong_wire = canonical.data.clone();
    encode_varint_into(&mut wrong_wire, (u64::from(21_u32) << 3) | 2);
    encode_varint_into(&mut wrong_wire, 1);
    wrong_wire.push(1);
    malformed_payloads.push(wrong_wire);

    let mut noncanonical_value = canonical.data.clone();
    encode_varint_into(&mut noncanonical_value, u64::from(20_u32) << 3);
    noncanonical_value.extend_from_slice(&[0x80, 0x00]);
    malformed_payloads.push(noncanonical_value);

    let mut zero_page = canonical.data;
    litchi_iwa_common::wire::append_varint_field(&mut zero_page, 22, 0)?;
    malformed_payloads.push(zero_page);

    for malformed in malformed_payloads {
        let malformed_bytes = synthetic_package_with_first_message(
            "Alpha",
            RawMessage {
                type_: SECTION_MESSAGE_TYPE,
                data: malformed,
            },
            "Beta",
            empty,
        )?;
        let malformed_package = Package::from_bytes(&malformed_bytes)?;
        assert!(matches!(
            malformed_package.section_pagination(SectionSelector::index(0)),
            Err(SectionPaginationError::InvalidSource)
        ));
        assert!(matches!(
            malformed_package.edit_section_pagination(SectionSelector::index(0)),
            Err(SectionPaginationError::InvalidSource)
        ));
        assert_eq!(malformed_package.source_bytes(), malformed_bytes);
    }
    Ok(())
}

#[test]
fn no_op_supports_legacy_but_changed_legacy_source_is_refused() -> TestResult<()> {
    let original = pagination(
        Some(Start::NextPage),
        Some(PageNumbering::ContinueFromPrevious),
        Some(1),
    )?;
    let flat = synthetic_package("Alpha", original, "Beta", Pagination::new())?;
    let legacy = legacy_package_bytes(&flat)?;
    let package = Package::from_bytes(&legacy)?;

    let mut noop_edit = package.edit_section_pagination(SectionSelector::index(0))?;
    noop_edit.set_pagination(original)?;
    let noop_commit = noop_edit.commit()?;
    assert!(noop_commit.patch().is_noop());
    assert_eq!(noop_commit.package().source_bytes(), legacy);
    assert_eq!(
        package
            .apply_section_pagination(noop_commit.patch())?
            .package()
            .source_bytes(),
        legacy
    );

    let mut changed = package.edit_section_pagination(SectionSelector::index(0))?;
    changed.set_start(Some(Start::LeftPage))?;
    assert!(matches!(
        changed.commit(),
        Err(SectionPaginationError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn changed_pagination_preserves_opaque_content_and_inverse_restores_source() -> TestResult<()> {
    let before = Pagination::new();
    let unchanged = pagination(
        Some(Start::RightPage),
        Some(PageNumbering::ContinueFromPrevious),
        Some(9),
    )?;
    let bytes = synthetic_package("Alpha", before, "Beta", unchanged)?;
    let package = Package::from_bytes(&bytes)?;
    let source_text = package.text()?;
    let source_first_payload =
        message_payload(&bytes, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let source_first_header = object_header(&bytes, FIRST_SECTION_IDENTIFIER)?;
    let source_second_payload =
        message_payload(&bytes, SECOND_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let source_body_payload = message_payload(&bytes, BODY_IDENTIFIER, 2_001)?;
    let after = pagination(
        Some(Start::LeftPage),
        Some(PageNumbering::Restart),
        Some(u32::MAX),
    )?;

    let mut edit = package.edit_section_pagination(SectionSelector::name("Alpha"))?;
    edit.set_pagination(after)?;
    let commit = edit.commit()?;
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().position(), Position::new(0));
    assert_eq!(commit.patch().before(), before);
    assert_eq!(commit.patch().after(), after);
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_eq!(commit.package().text()?, source_text);
    assert_eq!(
        commit
            .package()
            .section_pagination(SectionSelector::index(0))?,
        after
    );
    assert_eq!(
        commit
            .package()
            .section_pagination(SectionSelector::index(1))?,
        unchanged
    );

    let target = commit.package().source_bytes();
    let target_first_payload =
        message_payload(target, FIRST_SECTION_IDENTIFIER, SECTION_MESSAGE_TYPE)?;
    let target_first_header = object_header(target, FIRST_SECTION_IDENTIFIER)?;
    assert_eq!(
        fields_except_pagination(&target_first_payload)?,
        fields_except_pagination(&source_first_payload)?
    );
    assert!(
        WireView::parse(&target_first_payload)?
            .fields()
            .any(|field| field.number() == 99)
    );
    assert_eq!(
        normalize_message_length(&target_first_header, 1)?,
        normalize_message_length(&source_first_header, 1)?,
        "only the pagination payload length may change in the opaque archive header"
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

    let applied = package.apply_section_pagination(commit.patch())?;
    assert_eq!(applied.package().source_bytes(), target);
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), after);
    assert_eq!(inverse.after(), before);
    let restored = commit.package().apply_section_pagination(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(
        restored
            .package()
            .section_pagination(SectionSelector::index(0))?,
        before
    );

    let unrelated = Package::from_bytes(&synthetic_package("Other", before, "Beta", unchanged)?)?;
    assert!(matches!(
        unrelated.apply_section_pagination(commit.patch()),
        Err(SectionPaginationError::PatchConflict)
    ));
    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains(PRIVATE_NATIVE_MARKER));
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("fingerprint"));
    assert!(!debug.contains("bytes"));

    assert_send_sync(&package);
    assert_send_sync(&commit);
    assert_send_sync(commit.patch());
    assert_send_sync(commit.diagnostics());
    assert_type_send_sync::<SectionPaginationError>();
    Ok(())
}

#[test]
fn changed_pagination_respects_retained_output_limit() -> TestResult<()> {
    let before = Pagination::new();
    let bytes = synthetic_package("Alpha", before, "Beta", before)?;
    let target = pagination(
        Some(Start::Unknown(u32::MAX)),
        Some(PageNumbering::Unknown(u32::MAX)),
        Some(u32::MAX),
    )?;
    let unrestricted = Package::from_bytes(&bytes)?;
    let mut unrestricted_edit = unrestricted.edit_section_pagination(SectionSelector::index(0))?;
    unrestricted_edit.set_pagination(target)?;
    let target_length = unrestricted_edit.commit()?.package().source_bytes().len();
    assert!(target_length > bytes.len());

    let maximum = u64::try_from(target_length - 1)?;
    let limits = Limits::new(maximum, 8, 1_024 * 1_024, 1_024 * 1_024, 1_024 * 1_024)?;
    let package = Package::from_bytes_with_limits(&bytes, limits)?;
    let mut limited_edit = package.edit_section_pagination(SectionSelector::index(0))?;
    limited_edit.set_pagination(target)?;
    assert!(matches!(
        limited_edit.commit(),
        Err(SectionPaginationError::LimitExceeded {
            kind: SectionPaginationLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}
