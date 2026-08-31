//! Focused exact-source coverage for moving an existing Numbers table between
//! rooted sheets.
//!
//! A relocation changes only ownership edges: the table-info drawable moves
//! from one rooted sheet's ordered drawable list to another, and its parent
//! reference follows it.  The table model and every other package member are
//! intentionally treated as opaque content by this suite.

use std::{fmt::Debug, io};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{WireView, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tst};
use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    table::relocation::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const TABLES_MEMBER: &str = "Index/Tables.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const SENTINEL_MEMBER: &str = "Data/sentinel.bin";

const DOCUMENT_ID: u64 = 1;
const SOURCE_SHEET_ID: u64 = 2;
const DESTINATION_SHEET_ID: u64 = 50;
const ALPHA_INFO_ID: u64 = 10;
const BETA_INFO_ID: u64 = 11;
const GAMMA_INFO_ID: u64 = 12;
const ALPHA_MODEL_ID: u64 = 20;
const BETA_MODEL_ID: u64 = 21;
const GAMMA_MODEL_ID: u64 = 22;
const SIDECAR_ID: u64 = 90;

const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;

const SOURCE_SHEET_NAME: &str = "Source";
const DESTINATION_SHEET_NAME: &str = "Destination";
const ALPHA_TABLE_NAME: &str = "Alpha";
const BETA_TABLE_NAME: &str = "Beta";
const GAMMA_TABLE_NAME: &str = "Gamma";

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        bytes
    }
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

fn table_model(identifier: u64, name: &str) -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: format!("relocation-table-{identifier}"),
        table_name: name.to_owned(),
        number_of_rows: 2,
        number_of_columns: 2,
        base_data_store: tst::DataStore {
            string_table: reference(SIDECAR_ID),
            formula_table: reference(SIDECAR_ID),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn table_info(
    identifier: u64,
    model_identifier: u64,
    parent_identifier: u64,
    name: &str,
) -> TestResult<ArchiveObject> {
    let mut payload = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(reference(parent_identifier)),
            ..Default::default()
        },
        table_model: reference(model_identifier),
        ..Default::default()
    }
    .encode_to_vec();
    // This field is deliberately not modeled by the semantic owner.  A move
    // must retain it byte-for-byte even though the known parent reference is
    // rewritten.
    append_varint_field(&mut payload, 90, u64::from(name.len() as u32) + identifier)?;

    let mut result = object(identifier, TABLE_INFO_MESSAGE_TYPE, payload)?;
    result.archive_info.message_infos[0].object_references =
        vec![parent_identifier, model_identifier];
    let mut parent_path = FieldInfo::new(vec![1, 2]);
    parent_path.object_references = vec![parent_identifier];
    result.archive_info.message_infos[0].field_infos = vec![parent_path];
    Ok(result)
}

fn sidecars() -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        SIDECAR_ID,
        [
            tst::table_data_list::ListType::String,
            tst::table_data_list::ListType::Formula,
        ]
        .into_iter()
        .map(|list_type| RawMessage {
            type_: TABLE_DATA_LIST_MESSAGE_TYPE,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        })
        .collect(),
    )?)
}

fn sheet(
    identifier: u64,
    name: &str,
    drawable_identifiers: &[u64],
    unknown_field: u32,
) -> TestResult<ArchiveObject> {
    let mut payload = tn::SheetArchive {
        name: name.to_owned(),
        drawable_infos: drawable_identifiers
            .iter()
            .copied()
            .map(reference)
            .collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut payload, unknown_field, identifier + 900)?;

    let mut result = object(identifier, SHEET_MESSAGE_TYPE, payload)?;
    result.archive_info.message_infos[0].object_references = drawable_identifiers.to_vec();
    let mut drawable_path = FieldInfo::new(vec![2]);
    drawable_path.object_references = drawable_identifiers.to_vec();
    result.archive_info.message_infos[0].field_infos = vec![drawable_path];
    Ok(result)
}

/// Build a strict rooted graph with two source tables and one destination
/// table.  The selected table has a stable parent edge and the destination is
/// non-empty so append order is observable.
fn fixture() -> TestResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT_ID,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(SOURCE_SHEET_ID), reference(DESTINATION_SHEET_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references =
        vec![SOURCE_SHEET_ID, DESTINATION_SHEET_ID];

    let document_component = SnappyStream::compress(
        &Archive {
            objects: vec![
                document,
                sheet(
                    SOURCE_SHEET_ID,
                    SOURCE_SHEET_NAME,
                    &[ALPHA_INFO_ID, BETA_INFO_ID],
                    91,
                )?,
                sheet(
                    DESTINATION_SHEET_ID,
                    DESTINATION_SHEET_NAME,
                    &[GAMMA_INFO_ID],
                    92,
                )?,
            ],
        }
        .to_bytes()?,
    )?;

    let tables_component = SnappyStream::compress(
        &Archive {
            objects: vec![
                table_info(
                    ALPHA_INFO_ID,
                    ALPHA_MODEL_ID,
                    SOURCE_SHEET_ID,
                    ALPHA_TABLE_NAME,
                )?,
                table_info(
                    BETA_INFO_ID,
                    BETA_MODEL_ID,
                    SOURCE_SHEET_ID,
                    BETA_TABLE_NAME,
                )?,
                table_info(
                    GAMMA_INFO_ID,
                    GAMMA_MODEL_ID,
                    DESTINATION_SHEET_ID,
                    GAMMA_TABLE_NAME,
                )?,
                object(
                    ALPHA_MODEL_ID,
                    TABLE_MODEL_MESSAGE_TYPE,
                    table_model(ALPHA_MODEL_ID, ALPHA_TABLE_NAME).encode_to_vec(),
                )?,
                object(
                    BETA_MODEL_ID,
                    TABLE_MODEL_MESSAGE_TYPE,
                    table_model(BETA_MODEL_ID, BETA_TABLE_NAME).encode_to_vec(),
                )?,
                object(
                    GAMMA_MODEL_ID,
                    TABLE_MODEL_MESSAGE_TYPE,
                    table_model(GAMMA_MODEL_ID, GAMMA_TABLE_NAME).encode_to_vec(),
                )?,
                sidecars()?,
            ],
        }
        .to_bytes()?,
    )?;

    let unrelated_component = SnappyStream::compress(
        &Archive {
            objects: vec![ArchiveObject::new(
                900,
                vec![RawMessage {
                    type_: 99_999,
                    data: b"unrelated Numbers component".to_vec(),
                }],
            )?],
        }
        .to_bytes()?,
    )?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (TABLES_MEMBER, tables_component.as_slice()),
            (UNRELATED_MEMBER, unrelated_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn component_stream(source: &[u8], member_name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("missing component {member_name}")))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_component(
    source: &[u8],
    member_name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&component_stream(source, member_name)?)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[EntryEdit::new(member_name, &compressed)],
        Limits::default(),
    )?)
}

fn with_duplicate_owner(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, DOCUMENT_MEMBER, |archive| {
        let target = archive
            .object_mut(DESTINATION_SHEET_ID)
            .ok_or_else(|| io::Error::other("destination sheet is missing"))?;
        let message = target
            .messages
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("destination sheet payload is missing"))?;
        let mut decoded = tn::SheetArchive::decode(message.data.as_slice())?;
        decoded.drawable_infos.push(reference(BETA_INFO_ID));
        target.replace_message(
            0,
            RawMessage {
                type_: message.type_,
                data: decoded.encode_to_vec(),
            },
        )?;
        target.archive_info.message_infos[0]
            .object_references
            .push(BETA_INFO_ID);
        Ok(())
    })
}

fn with_parent_mismatch(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, TABLES_MEMBER, |archive| {
        let table_info = archive
            .object_mut(BETA_INFO_ID)
            .ok_or_else(|| io::Error::other("selected table-info is missing"))?;
        let message = table_info
            .messages
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("selected table-info payload is missing"))?;
        let mut decoded = tst::TableInfoArchive::decode(message.data.as_slice())?;
        decoded.super_.parent = Some(reference(DESTINATION_SHEET_ID));
        table_info.replace_message(
            0,
            RawMessage {
                type_: message.type_,
                data: decoded.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn with_malformed_table_info(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_component(source, TABLES_MEMBER, |archive| {
        let table_info = archive
            .object_mut(BETA_INFO_ID)
            .ok_or_else(|| io::Error::other("selected table-info is missing"))?;
        let message = table_info
            .messages
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("selected table-info payload is missing"))?;
        table_info.replace_message(
            0,
            RawMessage {
                type_: message.type_,
                // This is a truncated length-delimited protobuf field.  The
                // archive framing remains valid so ingress, rather than the
                // fixture builder, gets to reject the malformed graph.
                data: vec![0x0a, 0x80],
            },
        )?;
        Ok(())
    })
}

fn object_message(
    source: &[u8],
    member_name: &str,
    identifier: u64,
    message_type: u32,
) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&component_stream(source, member_name)?)?;
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == message_type)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| {
            io::Error::other(format!(
                "object {identifier} message {message_type} is missing"
            ))
            .into()
        })
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn sheet_drawable_ids(source: &[u8], sheet_identifier: u64) -> TestResult<Vec<u64>> {
    Ok(tn::SheetArchive::decode(
        object_message(
            source,
            DOCUMENT_MEMBER,
            sheet_identifier,
            SHEET_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?
    .drawable_infos
    .into_iter()
    .map(|drawable| drawable.identifier)
    .collect())
}

fn table_parent(source: &[u8], table_identifier: u64) -> TestResult<u64> {
    tst::TableInfoArchive::decode(
        object_message(
            source,
            TABLES_MEMBER,
            table_identifier,
            TABLE_INFO_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?
    .super_
    .parent
    .map(|parent| parent.identifier)
    .ok_or_else(|| io::Error::other("table-info parent reference is missing").into())
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let expected_changed = [DOCUMENT_MEMBER, TABLES_MEMBER];
    let mut changed = Vec::new();
    for before in source_catalog.iter() {
        let after = target_catalog
            .iter()
            .find(|candidate| candidate.name() == before.name())
            .ok_or_else(|| io::Error::other("relocation removed a source member"))?;
        if before.data() != after.data() {
            changed.push(before.name().to_owned());
        } else {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
                "unchanged member {} lost its exact local record",
                before.name()
            );
        }
    }
    assert_eq!(changed, expected_changed);
    assert_eq!(source_catalog.len(), target_catalog.len());
    Ok(())
}

fn table_names(package: &Package, sheet: usize) -> Vec<String> {
    package.sheets()[sheet]
        .tables()
        .map(|table| table.name().to_owned())
        .collect()
}

fn assert_atomic_refusal(source: &[u8], label: &str) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        // Strict ingress is allowed to reject an ambiguous rooted graph.
        return Ok(());
    };
    let before = package.exact_bytes();
    assert!(
        package
            .move_table(SOURCE_SHEET_NAME, BETA_TABLE_NAME, DESTINATION_SHEET_NAME)
            .is_err(),
        "malformed graph was accepted for relocation: {label}"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn relocation_api_values_are_send_sync_debug_and_errors_are_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<SheetSelector<'static>>();
    assert_send_sync_debug::<TableSelector<'static>>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let bytes = fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.exact_bytes();
    let error = package
        .move_table("Missing source", BETA_TABLE_NAME, DESTINATION_SHEET_NAME)
        .expect_err("an unknown source sheet must not stage a move");
    let rendered = format!("{error:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Tables.iwa"));
    assert!(!rendered.contains("Data/sentinel"));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn same_sheet_move_is_an_exact_noop_and_replays_without_writing() -> TestResult {
    let bytes = fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.exact_bytes();
    let commit = package.move_table(SOURCE_SHEET_NAME, BETA_TABLE_NAME, SOURCE_SHEET_NAME)?;
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.patch().path(),
        Path::Table {
            source_sheet: 0,
            table: 1,
            destination_sheet: 0,
        }
    );
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(commit.package().exact_bytes(), before);
    assert_eq!(
        table_names(commit.package(), 0),
        vec![ALPHA_TABLE_NAME.to_owned(), BETA_TABLE_NAME.to_owned()]
    );
    let replay = package.apply_table_move(commit.patch())?;
    assert!(replay.patch().is_noop());
    assert_eq!(replay.package().exact_bytes(), before);
    let indexed = package.move_table(
        SheetSelector::index(0),
        TableSelector::index(1),
        SheetSelector::index(0),
    )?;
    assert!(indexed.patch().is_noop());
    assert_eq!(indexed.package().exact_bytes(), before);
    Ok(())
}

#[test]
fn changed_move_appends_in_order_preserves_content_unknowns_and_locality() -> TestResult {
    let bytes = fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let beta = package.sheets()[0]
        .tables()
        .find(|table| table.name() == BETA_TABLE_NAME)
        .cloned()
        .ok_or_else(|| io::Error::other("selected table is missing"))?;
    let gamma = package.sheets()[1]
        .tables()
        .find(|table| table.name() == GAMMA_TABLE_NAME)
        .cloned()
        .ok_or_else(|| io::Error::other("destination table is missing"))?;
    let source_sheet =
        object_message(&bytes, DOCUMENT_MEMBER, SOURCE_SHEET_ID, SHEET_MESSAGE_TYPE)?;
    let destination_sheet = object_message(
        &bytes,
        DOCUMENT_MEMBER,
        DESTINATION_SHEET_ID,
        SHEET_MESSAGE_TYPE,
    )?;
    let table_info = object_message(&bytes, TABLES_MEMBER, BETA_INFO_ID, TABLE_INFO_MESSAGE_TYPE)?;
    let beta_model = object_message(
        &bytes,
        TABLES_MEMBER,
        BETA_MODEL_ID,
        TABLE_MODEL_MESSAGE_TYPE,
    )?;
    let source_unknown = raw_fields(&source_sheet, 91)?;
    let destination_unknown = raw_fields(&destination_sheet, 92)?;
    let table_unknown = raw_fields(&table_info, 90)?;

    let commit = package.move_table(SOURCE_SHEET_NAME, BETA_TABLE_NAME, DESTINATION_SHEET_NAME)?;
    let target = commit.package().exact_bytes();
    let reopened = Package::from_bytes(&target)?;
    assert!(!commit.patch().is_noop());
    assert_eq!(
        commit.patch().path(),
        Path::Table {
            source_sheet: 0,
            table: 1,
            destination_sheet: 1,
        }
    );
    assert_eq!(commit.patch().source_sheet_position(), 0);
    assert_eq!(commit.patch().table_position(), 1);
    assert_eq!(commit.patch().destination_sheet_position(), 1);
    assert_eq!(commit.patch().destination_table_position(), 1);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        table_names(commit.package(), 0),
        vec![ALPHA_TABLE_NAME.to_owned()]
    );
    assert_eq!(
        table_names(commit.package(), 1),
        vec![GAMMA_TABLE_NAME.to_owned(), BETA_TABLE_NAME.to_owned()]
    );
    assert_eq!(
        sheet_drawable_ids(&target, SOURCE_SHEET_ID)?,
        vec![ALPHA_INFO_ID]
    );
    assert_eq!(
        sheet_drawable_ids(&target, DESTINATION_SHEET_ID)?,
        vec![GAMMA_INFO_ID, BETA_INFO_ID]
    );
    assert_eq!(table_parent(&bytes, BETA_INFO_ID)?, SOURCE_SHEET_ID);
    assert_eq!(table_parent(&target, BETA_INFO_ID)?, DESTINATION_SHEET_ID);
    assert_eq!(
        table_names(&reopened, 1),
        vec![GAMMA_TABLE_NAME.to_owned(), BETA_TABLE_NAME.to_owned()]
    );
    assert_eq!(
        commit.package().sheets()[1].tables().nth(0).cloned(),
        Some(gamma)
    );
    assert_eq!(
        commit.package().sheets()[1].tables().nth(1).cloned(),
        Some(beta)
    );
    assert_eq!(
        raw_fields(
            &object_message(
                &target,
                DOCUMENT_MEMBER,
                SOURCE_SHEET_ID,
                SHEET_MESSAGE_TYPE
            )?,
            91
        )?,
        source_unknown
    );
    assert_eq!(
        raw_fields(
            &object_message(
                &target,
                DOCUMENT_MEMBER,
                DESTINATION_SHEET_ID,
                SHEET_MESSAGE_TYPE
            )?,
            92
        )?,
        destination_unknown
    );
    assert_eq!(
        raw_fields(
            &object_message(
                &target,
                TABLES_MEMBER,
                BETA_INFO_ID,
                TABLE_INFO_MESSAGE_TYPE
            )?,
            90
        )?,
        table_unknown
    );
    assert_eq!(
        object_message(
            &target,
            TABLES_MEMBER,
            BETA_MODEL_ID,
            TABLE_MODEL_MESSAGE_TYPE
        )?,
        beta_model
    );
    assert_locality(&bytes, &target)?;
    let source_catalog = Catalog::from_bytes(&bytes)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    let source_sentinel = source_catalog
        .iter()
        .find(|entry| entry.name() == SENTINEL_MEMBER)
        .ok_or_else(|| io::Error::other("source sentinel is missing"))?;
    let target_sentinel = target_catalog
        .iter()
        .find(|entry| entry.name() == SENTINEL_MEMBER)
        .ok_or_else(|| io::Error::other("target sentinel is missing"))?;
    assert_eq!(source_sentinel.data(), target_sentinel.data());
    Ok(())
}

#[test]
fn selectors_fail_atomically_for_missing_source_table_and_destination() -> TestResult {
    let bytes = fixture()?;
    let package = Package::from_bytes(&bytes)?;
    for (source, table, destination) in [
        ("Missing source", BETA_TABLE_NAME, DESTINATION_SHEET_NAME),
        (SOURCE_SHEET_NAME, "Missing table", DESTINATION_SHEET_NAME),
        (SOURCE_SHEET_NAME, BETA_TABLE_NAME, "Missing destination"),
    ] {
        let before = package.exact_bytes();
        assert!(package.move_table(source, table, destination).is_err());
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn malformed_or_ambiguous_ownership_is_refused_without_partial_mutation() -> TestResult {
    let bytes = fixture()?;
    assert_atomic_refusal(&with_duplicate_owner(&bytes)?, "duplicate rooted owner")?;
    assert_atomic_refusal(&with_parent_mismatch(&bytes)?, "parent/owner mismatch")?;
    assert_atomic_refusal(&with_malformed_table_info(&bytes)?, "malformed table-info")?;
    Ok(())
}

#[test]
fn changed_patch_is_exact_source_bound_and_inverse_restores_every_byte() -> TestResult {
    let bytes = fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package.move_table(SOURCE_SHEET_NAME, BETA_TABLE_NAME, DESTINATION_SHEET_NAME)?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        package
            .apply_table_move(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_move(&inverse)?
            .package()
            .exact_bytes(),
        bytes
    );

    let tampered = Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
        &[EntryEdit::new(
            SENTINEL_MEMBER,
            b"tampered unrelated sentinel",
        )],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered)?;
    let tampered_before = tampered.exact_bytes();
    assert!(tampered.apply_table_move(commit.patch()).is_err());
    assert_eq!(tampered.exact_bytes(), tampered_before);
    assert!(package.apply_table_move(&inverse).is_err());
    Ok(())
}
