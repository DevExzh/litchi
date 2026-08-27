//! Exact-source integration coverage for Pages body-table names.
//!
//! The neighboring header fixture supplies the rooted body/table graph.  The
//! tests in this file deliberately keep the value side archive-free while
//! exercising the name owner at each package boundary: selector resolution,
//! strict wire admission, authority checks, bounded rewrite, reopen, and
//! reversible publication.

use std::collections::BTreeMap;
use std::error::Error as StdError;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{Archive, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::tsp;
use litchi_pages::{
    BodyTableName, BodyTableNameError as Error, BodyTableNameValueError, BodyTableSelector,
    Limits as PackageLimits, Package,
};
use prost::Message as _;

#[path = "body_table_header_settings.rs"]
mod header_fixture;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT_ID: u64 = 50_000;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const FIRST_DRAWABLE_IDENTIFIER: u64 = 200;
const FIRST_MODEL_IDENTIFIER: u64 = 300;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

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

fn source_package(names: [&str; 2]) -> TestResult<Vec<u8>> {
    header_fixture::synthetic_package(names, None)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = document_archive(package)?;
    let object = archive.object(identifier).ok_or("missing table model")?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table-model message")?
        .data
        .clone())
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    header_fixture::rewrite_document_archive(package, mutate)
}

fn append_selected_model_raw(package: &[u8], raw: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        let index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        let mut data = model.messages[index].data.clone();
        data.extend_from_slice(raw);
        model.replace_message_preserving_header(
            index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn append_model_role_alias(package: &[u8], alias_type: u32) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        let index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        let data = model.messages[index].data.clone();
        let mut info = model.archive_info.message_infos[index].clone();
        info.type_ = alias_type;
        info.length = u32::try_from(data.len())?;
        model.messages.push(RawMessage {
            type_: alias_type,
            data,
        });
        model.archive_info.message_infos.push(info);
        Ok(())
    })
}

fn duplicate_canonical_model_role(package: &[u8]) -> TestResult<Vec<u8>> {
    append_model_role_alias(package, TABLE_MODEL_MESSAGE_TYPE)
}

fn append_table_info_model_role_alias(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let table_info = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table info")?;
        let index = table_info
            .messages
            .iter()
            .position(|message| message.type_ == LEGACY_TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        let data = table_info.messages[index].data.clone();
        let mut info = table_info.archive_info.message_infos[index].clone();
        info.type_ = TABLE_MODEL_MESSAGE_TYPE;
        info.length = u32::try_from(data.len())?;
        table_info.messages.push(RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPE,
            data,
        });
        table_info.archive_info.message_infos.push(info);
        Ok(())
    })
}

fn append_archive_info_reference(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .object_references
            .push(identifier);
        Ok(())
    })
}

fn append_unknown_model_field(package: &[u8]) -> TestResult<Vec<u8>> {
    let mut raw = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut raw, 97, 0xdecafbad)?;
    append_selected_model_raw(package, &raw)
}

fn append_unknown_model_group(package: &[u8]) -> TestResult<Vec<u8>> {
    append_selected_model_raw(package, &[0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06])
}

fn model_field_varint(package: &[u8], field_number: u32) -> TestResult<u64> {
    let payload = model_payload(package, FIRST_MODEL_IDENTIFIER)?;
    let field = WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == field_number)
        .ok_or("missing model field")?;
    Ok(decode_varint_from_bytes(field.payload())?.0)
}

fn member_bytes(package: &[u8]) -> TestResult<BTreeMap<String, Vec<u8>>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect())
}

fn without_previews(package: &[u8]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(package)?.reassemble_with_deletions_to_bytes(
            &[],
            &PREVIEWS,
            Limits::default(),
        )?,
    )
}

fn sentinel(package: &[u8]) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == "Data/sentinel.bin")
        .ok_or("missing sentinel")?
        .data()
        .to_vec())
}

fn with_foreign_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let inbound =
        header_fixture::object(7_000, 7_001, vec![0x08, 0x01], &[FIRST_MODEL_IDENTIFIER])?;
    let foreign = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members: Vec<(String, Vec<u8>)> = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect();
    members.push(("Index/Foreign.iwa".to_owned(), foreign.to_vec()));
    let refs = members
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        refs,
        Limits::default(),
    )?)
}

fn with_foreign_field_data_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut inbound = header_fixture::object(7_000, 7_001, vec![0x08, 0x01], &[])?;
    let mut field = FieldInfo::new(FieldPath::from(vec![1]));
    field.data_references.push(FIRST_MODEL_IDENTIFIER);
    inbound.archive_info.message_infos[0]
        .field_infos
        .push(field);
    let foreign = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members: Vec<(String, Vec<u8>)> = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect();
    members.push(("Index/Foreign.iwa".to_owned(), foreign.to_vec()));
    let refs = members
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        refs,
        Limits::default(),
    )?)
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn metadata_package(source: &[u8]) -> TestResult<Vec<u8>> {
    let source_catalog = Catalog::from_bytes(source)?;
    let document = source_catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let archive = Archive::parse(SnappyStream::decompress(document.data())?.as_bytes())?;
    let mut identifiers = archive
        .objects
        .iter()
        .map(|object| {
            object
                .archive_info
                .identifier
                .ok_or("missing object identifier")
        })
        .collect::<Result<Vec<_>, _>>()?;
    identifiers.push(METADATA_OBJECT_ID);
    identifiers.sort_unstable();
    let metadata = tsp::PackageMetadata {
        last_object_identifier: METADATA_OBJECT_ID,
        save_token: Some(1),
        components: vec![tsp::ComponentInfo {
            identifier: 1,
            preferred_locator: "Document".to_owned(),
            locator: Some("Document".to_owned()),
            save_token: Some(1),
            object_uuid_map_entries: identifiers.iter().copied().map(uuid_entry).collect(),
            ..tsp::ComponentInfo::default()
        }],
        ..tsp::PackageMetadata::default()
    }
    .encode_to_vec();
    let metadata_archive = SnappyStream::compress(
        &Archive {
            objects: vec![header_fixture::object(
                METADATA_OBJECT_ID,
                METADATA_MESSAGE_TYPE,
                metadata,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;
    let mut members = source_catalog
        .iter()
        .filter(|entry| entry.name() != METADATA_MEMBER)
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    members.push((METADATA_MEMBER.to_owned(), metadata_archive.to_vec()));
    let refs = members
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        refs,
        Limits::default(),
    )?)
}

fn rewrite_metadata(
    source: &[u8],
    mutate: impl FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing metadata member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let metadata_object = archive
        .object_mut(METADATA_OBJECT_ID)
        .ok_or("missing metadata object")?;
    let index = metadata_object
        .messages
        .iter()
        .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing metadata message")?;
    let mut metadata =
        tsp::PackageMetadata::decode(metadata_object.messages[index].data.as_slice())?;
    mutate(&mut metadata)?;
    metadata_object.replace_message_preserving_header(
        index,
        RawMessage {
            type_: METADATA_MESSAGE_TYPE,
            data: metadata.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(METADATA_MEMBER, compressed.as_slice())],
        Limits::default(),
    )?)
}

fn append_metadata_unknown(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing metadata member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let metadata_object = archive
        .object_mut(METADATA_OBJECT_ID)
        .ok_or("missing metadata object")?;
    let index = metadata_object
        .messages
        .iter()
        .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing metadata message")?;
    let mut data = metadata_object.messages[index].data.clone();
    litchi_iwa_common::wire::append_varint_field(&mut data, 90, 0xdecafbad)?;
    metadata_object.replace_message_preserving_header(
        index,
        RawMessage {
            type_: METADATA_MESSAGE_TYPE,
            data,
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(METADATA_MEMBER, compressed.as_slice())],
        Limits::default(),
    )?)
}

fn assert_rejected(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        Catalog::from_bytes(source)?;
        return Ok(());
    };
    let before = package.exact_bytes();
    assert!(package.body_table_name(0usize).is_err());
    if let Ok(edit) = package.edit_body_table_name(0usize) {
        let edit = edit
            .set_name("replacement")
            .expect("a valid replacement name must be accepted");
        assert!(edit.commit().is_err());
    }
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn selectors_read_current_name_and_reject_ambiguity() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let by_position = package.body_table_name(BodyTableSelector::index(0))?;
    let by_name = package.body_table_name(BodyTableSelector::name("Revenue"))?;
    assert_eq!(by_position.as_str(), "Revenue");
    assert_eq!(by_position, by_name);
    assert!(matches!(
        package.body_table_name(BodyTableSelector::name("Missing")),
        Err(Error::TableNotFound)
    ));

    let ambiguous = Package::from_bytes(&source_package(["Revenue", "Revenue"])?)?;
    assert!(matches!(
        ambiguous.body_table_name(BodyTableSelector::name("Revenue")),
        Err(Error::AmbiguousTableName | Error::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn no_op_is_exact_and_patch_is_reversible() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_name(0usize)?;
    let commit = package
        .edit_body_table_name(BodyTableSelector::name("Revenue"))?
        .set(before.clone())
        .commit()?;
    assert_eq!(commit.patch().before(), &before);
    assert_eq!(commit.patch().after(), &before);
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(commit.package().exact_bytes(), source.as_slice());
    let applied = commit.package().apply_body_table_name(commit.patch())?;
    assert_eq!(applied.package().exact_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn changed_name_reopens_preserves_unknowns_and_inverts_exactly() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let new_name = BodyTableName::new("Revenu\u{e9} \u{1f680}")?;
    let commit = package
        .edit_body_table_name(0usize)?
        .set(new_name.clone())
        .commit()?;
    assert_eq!(
        commit.package().body_table_name(0usize)?.as_str(),
        new_name.as_str()
    );
    assert_eq!(commit.patch().after(), &new_name);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        sentinel(&commit.package().exact_bytes())?,
        b"untouched-sentinel"
    );
    assert_eq!(
        model_field_varint(&commit.package().exact_bytes(), UNKNOWN_MODEL_FIELD)?,
        UNKNOWN_MODEL_VALUE
    );
    let members_before = member_bytes(&source)?;
    let members_after = member_bytes(&commit.package().exact_bytes())?;
    assert_ne!(
        members_after.get(DOCUMENT_MEMBER),
        members_before.get(DOCUMENT_MEMBER)
    );
    assert!(
        PREVIEWS
            .iter()
            .all(|preview| !members_after.contains_key(*preview))
    );
    assert_eq!(
        members_after.get("Data/sentinel.bin"),
        members_before.get("Data/sentinel.bin")
    );
    let restored = commit
        .package()
        .apply_body_table_name(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    assert_eq!(
        restored.package().body_table_name(0usize)?.as_str(),
        "Revenue"
    );
    Ok(())
}

#[test]
fn empty_utf8_and_bounded_name_semantics_are_explicit() -> TestResult {
    assert!(matches!(
        BodyTableName::new(""),
        Err(BodyTableNameValueError::Empty)
    ));
    assert!(matches!(
        BodyTableName::new("bad\0name"),
        Err(BodyTableNameValueError::ContainsNul)
    ));
    let unicode = BodyTableName::new("收入表 — Café №42")?;
    assert_eq!(unicode.as_str(), "收入表 — Café №42");

    let source = without_previews(&source_package(["Revenue", "Costs"])?)?;
    let replacement = "x".repeat(1_024);
    let unrestricted = Package::from_bytes(&source)?
        .edit_body_table_name(0usize)?
        .set_name(&replacement)?
        .commit()?;
    assert!(unrestricted.package().exact_bytes().len() > source.len());
    let limits = PackageLimits::new(
        u64::try_from(source.len())?,
        32,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let package = Package::from_bytes_with_limits(&source, limits)?;
    let before = package.exact_bytes();
    let error = package
        .edit_body_table_name(0usize)?
        .set_name(&replacement)?
        .commit()
        .expect_err("output growth must be rejected at the physical ceiling");
    assert!(matches!(error, Error::LimitExceeded { .. }));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn locked_table_allows_exact_noop_but_refuses_change() -> TestResult {
    let source = header_fixture::synthetic_package(["Revenue", "Costs"], Some(true))?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_name(0usize)?;
    let noop = package
        .edit_body_table_name(0usize)?
        .set(before.clone())
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), source.as_slice());
    let changed = package
        .edit_body_table_name(0usize)?
        .set_name("locked replacement")?
        .commit();
    assert!(
        changed.is_err(),
        "a locked table must reject a changed name"
    );
    assert_eq!(package.exact_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn malformed_name_wire_is_rejected_without_mutation() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let cases: [&[u8]; 6] = [
        &[0x42, 0x01, b"X"[0]],    // duplicate field 8
        &[0x40, 0x01],             // field 8 with varint wire
        &[0x42, 0x81, 0x00, b'X'], // noncanonical length
        &[0x42],                   // truncated length-delimited field
        &[0xa3, 0x06, 0x08, 0x01], // unterminated unknown group
        &[0x42, 0x02, 0xff, 0xff], // invalid UTF-8
    ];
    for raw in cases {
        let malformed_source = append_selected_model_raw(&source, raw)?;
        assert_rejected(&malformed_source)?;
    }
    Ok(())
}

#[test]
fn unknown_fields_and_unknown_groups_survive_name_rewrite() -> TestResult {
    let source = append_unknown_model_group(&append_unknown_model_field(&source_package([
        "Revenue", "Costs",
    ])?)?)?;
    let package = Package::from_bytes(&source)?;
    let edit = package.edit_body_table_name(0usize)?;
    let edit = edit.set_name("Preserved unknowns")?;
    let commit = edit.commit()?;
    let changed_payload = model_payload(&commit.package().exact_bytes(), FIRST_MODEL_IDENTIFIER)?;
    let mut unknown_97 = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut unknown_97, 97, 0xdecafbad)?;
    let mut unknown_99 = Vec::new();
    litchi_iwa_common::wire::append_varint_field(
        &mut unknown_99,
        UNKNOWN_MODEL_FIELD,
        UNKNOWN_MODEL_VALUE,
    )?;
    assert!(
        changed_payload
            .windows(unknown_97.len())
            .any(|window| window == unknown_97.as_slice())
    );
    assert!(
        changed_payload
            .windows(unknown_99.len())
            .any(|window| window == unknown_99.as_slice())
    );
    assert!(
        changed_payload
            .windows(6)
            .any(|window| window == [0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06])
    );
    Ok(())
}

#[test]
fn canonical_and_legacy_role_aliases_fail_closed() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    for malformed in [
        append_model_role_alias(&source, LEGACY_TABLE_MODEL_MESSAGE_TYPE)?,
        duplicate_canonical_model_role(&source)?,
        append_model_role_alias(&source, TABLE_STYLE_MESSAGE_TYPE)?,
        append_table_info_model_role_alias(&source)?,
    ] {
        assert_rejected(&malformed)?;
    }
    Ok(())
}

#[test]
fn archive_info_and_global_inbound_authority_fail_closed() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    for malformed in [
        append_archive_info_reference(&source, 999_999)?,
        with_foreign_inbound(&source)?,
        with_foreign_field_data_inbound(&source)?,
    ] {
        assert_rejected(&malformed)?;
    }
    Ok(())
}

#[test]
fn metadata_uuid_and_unknown_authority_fail_closed() -> TestResult {
    let source = metadata_package(&source_package(["Revenue", "Costs"])?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.body_table_name(0usize)?.as_str(), "Revenue");
    let hostile = [
        rewrite_metadata(&source, |metadata| {
            metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?
                .object_uuid_map_entries
                .push(uuid_entry(FIRST_MODEL_IDENTIFIER));
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != FIRST_MODEL_IDENTIFIER);
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata.versioned_components.push(tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                object_uuid_map_entries: vec![uuid_entry(FIRST_MODEL_IDENTIFIER)],
                ..tsp::ComponentInfo::default()
            });
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?
                .external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: 1,
                    object_identifier: Some(FIRST_MODEL_IDENTIFIER),
                    is_weak: Some(false),
                });
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?
                .data_references
                .push(tsp::ComponentDataReference {
                    data_identifier: FIRST_MODEL_IDENTIFIER,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: FIRST_MODEL_IDENTIFIER,
                        count: 1,
                    }],
                });
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata
                .components
                .first_mut()
                .ok_or("missing metadata component")?
                .ambiguous_object_identifiers
                .push(FIRST_MODEL_IDENTIFIER);
            Ok(())
        })?,
        rewrite_metadata(&source, |metadata| {
            metadata.data_metadata_map = Some(tsp::Reference {
                identifier: FIRST_MODEL_IDENTIFIER,
                ..tsp::Reference::default()
            });
            Ok(())
        })?,
        append_metadata_unknown(&source)?,
    ];
    for malformed in hostile {
        assert_rejected(&malformed)?;
    }
    Ok(())
}

#[test]
fn rename_to_existing_sibling_name_is_rejected_atomically() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let result = package
        .edit_body_table_name(0usize)?
        .set_name("Costs")?
        .commit();
    assert!(matches!(
        result,
        Err(Error::AmbiguousTableName | Error::InvalidSource)
    ));
    assert_eq!(package.exact_bytes(), source);
    Ok(())
}

#[test]
fn stale_patch_conflict_is_source_bound() -> TestResult {
    let source = source_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_name(0usize)?
        .set_name("Candidate")?
        .commit()?;
    let catalog = Catalog::from_bytes(&source)?;
    let tampered_source = catalog.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_name(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(tampered.exact_bytes(), tampered_source.as_slice());
    Ok(())
}

#[test]
fn public_name_transaction_values_are_send_sync_debug_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<BodyTableName>();
    assert_send_sync_debug::<litchi_pages::BodyTableNameEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableNamePatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableNameCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableNameDiagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<litchi_pages::BodyTableNameLimitKind>();
    let package = Package::from_bytes(&source_package(["Revenue", "Costs"])?)?;
    let edit = package.edit_body_table_name(0usize)?;
    let debug = format!("{edit:?}");
    assert!(debug.contains("name"));
    assert!(!debug.contains("Index/Document.iwa"));
    assert!(!debug.contains("300"));
    Ok(())
}
