//! Native integration coverage for selector-first Pages body-table headers.
//!
//! The fixture deliberately keeps all seven header/footer fields in the table
//! model so that the semantic value can prove optional-field presence.  The
//! transaction tests also keep the table-info/model ownership graph populated;
//! a header edit must preserve that graph and the unrelated package members.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsd, tsp, tst, tswp};
use litchi_pages::table::headers::{Count, Settings};
use litchi_pages::{BodyTableHeaderSettingsError as Error, BodyTableSelector, Limits, Package};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_ATTACHMENT_IDENTIFIER: u64 = 100;
const SECOND_ATTACHMENT_IDENTIFIER: u64 = 110;
const FIRST_DRAWABLE_IDENTIFIER: u64 = 200;
const SECOND_DRAWABLE_IDENTIFIER: u64 = 210;
const FIRST_MODEL_IDENTIFIER: u64 = 300;
const SECOND_MODEL_IDENTIFIER: u64 = 310;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
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

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, identifier: u64) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.push(identifier);
    field
}

pub(crate) fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn fixture_settings() -> Settings {
    Settings {
        header_rows: Some(Count::TWO),
        header_columns: Some(Count::ONE),
        footer_rows: Some(Count::THREE),
        header_rows_frozen: Some(false),
        header_columns_frozen: Some(true),
        repeating_header_rows_enabled: Some(true),
        repeating_header_columns_enabled: Some(false),
    }
}

fn table_model(name: &str, settings: Settings) -> Vec<u8> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 8,
        number_of_columns: 6,
        number_of_header_rows: settings.header_rows.map(|count| count.get() as u32),
        number_of_header_columns: settings.header_columns.map(|count| count.get() as u32),
        number_of_footer_rows: settings.footer_rows.map(|count| count.get() as u32),
        header_rows_frozen: settings.header_rows_frozen,
        header_columns_frozen: settings.header_columns_frozen,
        repeating_header_rows_enabled: settings.repeating_header_rows_enabled,
        repeating_header_columns_enabled: settings.repeating_header_columns_enabled,
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut payload,
        UNKNOWN_MODEL_FIELD,
        UNKNOWN_MODEL_VALUE,
    )
    .expect("synthetic unknown header field fits");
    payload
}

pub(crate) fn synthetic_package(names: [&str; 2], locked: Option<bool>) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let entries = names
        .iter()
        .enumerate()
        .map(|(index, _)| {
            let attachment = if index == 0 {
                FIRST_ATTACHMENT_IDENTIFIER
            } else {
                SECOND_ATTACHMENT_IDENTIFIER
            };
            Ok(tswp::object_attribute_table::ObjectAttribute {
                character_index: u32::try_from(index)?,
                object: Some(reference(attachment)),
            })
        })
        .collect::<Result<Vec<_>, std::num::TryFromIntError>>()?;
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["\u{fffc}\u{fffc}".to_owned()],
        table_attachment: Some(tswp::ObjectAttributeTable { entries }),
        ..tswp::StorageArchive::default()
    };
    let attachment_ids = [FIRST_ATTACHMENT_IDENTIFIER, SECOND_ATTACHMENT_IDENTIFIER];
    let mut body_object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body.encode_to_vec(),
        &attachment_ids,
    )?;
    for attachment in attachment_ids {
        body_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![9], attachment));
    }
    let mut root_object = object(
        1,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));

    let mut objects = vec![root_object, body_object];
    for (index, name) in names.iter().enumerate() {
        let (attachment_identifier, drawable_identifier, model_identifier) = if index == 0 {
            (
                FIRST_ATTACHMENT_IDENTIFIER,
                FIRST_DRAWABLE_IDENTIFIER,
                FIRST_MODEL_IDENTIFIER,
            )
        } else {
            (
                SECOND_ATTACHMENT_IDENTIFIER,
                SECOND_DRAWABLE_IDENTIFIER,
                SECOND_MODEL_IDENTIFIER,
            )
        };
        let attachment_payload = tswp::DrawableAttachmentArchive {
            drawable: Some(reference(drawable_identifier)),
            ..tswp::DrawableAttachmentArchive::default()
        }
        .encode_to_vec();
        let mut table_info_payload = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(BODY_IDENTIFIER)),
                locked,
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(model_identifier),
            ..tst::TableInfoArchive::default()
        }
        .encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(&mut table_info_payload, 98, 17)?;

        let attachment_object = object(
            attachment_identifier,
            ATTACHMENT_MESSAGE_TYPE,
            attachment_payload,
            &[drawable_identifier],
        )?;
        let mut drawable_object = object(
            drawable_identifier,
            TABLE_INFO_MESSAGE_TYPE,
            table_info_payload,
            &[BODY_IDENTIFIER, model_identifier],
        )?;
        drawable_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![1, 2], BODY_IDENTIFIER),
                field_reference(vec![2], model_identifier),
            ]);
        let model_object = object(
            model_identifier,
            TABLE_MODEL_MESSAGE_TYPE,
            table_model(name, fixture_settings()),
            &[],
        )?;
        objects.extend([attachment_object, drawable_object, model_object]);
    }
    let archive = Archive { objects };
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    let mut members: Vec<(&str, &[u8])> = vec![
        ("Data/sentinel.bin", b"untouched-sentinel"),
        (DOCUMENT_MEMBER, component.as_slice()),
    ];
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name, b"preview".as_slice())),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        members,
        Limits::default(),
    )?)
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
    let object = archive.object(identifier).ok_or("missing model object")?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing model message")?
        .data
        .clone())
}

pub(crate) fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn append_selected_model_raw(package: &[u8], raw: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model")?;
        let message_index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut data = model.messages[message_index].data.clone();
        data.extend_from_slice(raw);
        model.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn append_dependency_field(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model")?;
        let message_index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut data = model.messages[message_index].data.clone();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut data,
            83,
            &reference(999).encode_to_vec(),
        )?;
        model.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        model.archive_info.message_infos[0]
            .object_references
            .push(999);
        Ok(())
    })
}

fn append_table_info_raw(package: &[u8], raw: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let drawable = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table-info object")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        let mut data = drawable.messages[message_index].data.clone();
        data.extend_from_slice(raw);
        drawable.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn make_table_model_reference_external(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let drawable = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table-info object")?;
        let message_index = drawable
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        let mut info =
            tst::TableInfoArchive::decode(drawable.messages[message_index].data.as_slice())?;
        info.table_model.deprecated_is_external = Some(true);
        drawable.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: info.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn sentinel(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == "Data/sentinel.bin")
        .ok_or("missing sentinel")?
        .data()
        .to_vec())
}

fn changed_settings() -> Settings {
    Settings {
        header_rows: None,
        header_columns: Some(Count::THREE),
        footer_rows: Some(Count::ONE),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(false),
        repeating_header_rows_enabled: Some(false),
        repeating_header_columns_enabled: None,
    }
}

#[test]
fn selectors_read_position_and_name_and_reject_ambiguity() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let package = Package::from_bytes(&source)?;
    let by_position = package.body_table_header_settings(BodyTableSelector::index(0))?;
    let by_name = package.body_table_header_settings(BodyTableSelector::name("Revenue"))?;
    assert_eq!(by_position, by_name);
    assert_eq!(by_position, fixture_settings());
    assert!(matches!(
        package.body_table_header_settings(BodyTableSelector::name("Missing")),
        Err(Error::TableNotFound)
    ));

    let ambiguous = Package::from_bytes(&synthetic_package(["Revenue", "Revenue"], None)?)?;
    assert!(matches!(
        ambiguous.body_table_header_settings(BodyTableSelector::name("Revenue")),
        Err(Error::AmbiguousTableName | Error::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn all_optional_fields_preserve_presence_and_effective_values() -> TestResult {
    assert!(Count::new(0).is_err());
    assert!(Count::new(6).is_err());
    assert_eq!(Count::new(1)?, Count::ONE);
    let package = Package::from_bytes(&synthetic_package(["Revenue", "Costs"], None)?)?;
    let before = package.body_table_header_settings(0usize)?;
    assert_eq!(before.header_row_count(), 2);
    assert_eq!(before.header_column_count(), 1);
    assert_eq!(before.footer_row_count(), 3);
    assert!(!before.header_rows_are_frozen());
    assert!(before.header_columns_are_frozen());
    assert!(before.repeats_header_rows());
    assert!(!before.repeats_header_columns());
    let after = changed_settings();
    assert_eq!(after.header_rows, None);
    assert_eq!(after.header_rows_frozen, Some(true));
    assert_eq!(after.repeating_header_columns_enabled, None);
    let commit = package
        .edit_body_table_header_settings(0usize)?
        .set(after)
        .commit()?;
    assert_eq!(commit.package().body_table_header_settings(0usize)?, after);
    Ok(())
}

#[test]
fn no_op_is_exact_and_presence_preserving() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_header_settings(0usize)?;
    let commit = package
        .edit_body_table_header_settings(BodyTableSelector::name("Revenue"))?
        .set(before)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(commit.package().exact_bytes(), source.as_slice());
    assert_eq!(
        commit
            .package()
            .apply_body_table_header_settings(commit.patch())?
            .package()
            .exact_bytes(),
        source.as_slice()
    );
    Ok(())
}

#[test]
fn changed_settings_preserve_unknowns_delete_previews_and_inverse_exactly() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let package = Package::from_bytes(&source)?;
    let after = changed_settings();
    let commit = package
        .edit_body_table_header_settings(0usize)?
        .set(after)
        .commit()?;
    assert_eq!(
        commit.package().body_table_header_settings("Revenue")?,
        after
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        sentinel(&commit.package().exact_bytes())?,
        b"untouched-sentinel"
    );
    let changed_model = model_payload(&commit.package().exact_bytes(), FIRST_MODEL_IDENTIFIER)?;
    let unknown = WireView::parse(&changed_model)?
        .fields()
        .find(|field| field.number() == UNKNOWN_MODEL_FIELD)
        .ok_or("unknown model field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(unknown.payload())?.0,
        UNKNOWN_MODEL_VALUE
    );
    let restored = commit
        .package()
        .apply_body_table_header_settings(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source.as_slice());
    assert_eq!(
        restored.package().body_table_header_settings(0usize)?,
        fixture_settings()
    );
    Ok(())
}

#[test]
fn patch_conflict_and_malformed_selected_fields_are_atomic() -> TestResult<()> {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_header_settings(0usize)?
        .set(changed_settings())
        .commit()?;
    let catalog = Catalog::from_bytes(&source)?;
    let tampered_source = catalog.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_header_settings(commit.patch()),
        Err(Error::PatchConflict)
    ));

    for raw in [
        &[0x48, 0][..],                            // header rows = zero
        &[0x50, 7][..],                            // header columns exceed the model
        &[0x58, 0][..],                            // footer rows = zero
        &[0x60, 2][..],                            // bool is neither false nor true
        &[0x48, 1, 0x48, 2][..],                   // duplicate header rows
        &[0x4d, 0, 0, 0, 0][..],                   // selected field has wrong wire type
        &[0x48, 0x80, 0][..],                      // selected value is non-canonical
        &[0xa3, 0x06, 0x08, 0x01, 0xa4, 0x06][..], // package ingress rejects groups
    ] {
        let malformed_source = append_selected_model_raw(&source, raw)?;
        let Ok(malformed) = Package::from_bytes(&malformed_source) else {
            Catalog::from_bytes(&malformed_source)?;
            continue;
        };
        let before = malformed.exact_bytes();
        assert!(matches!(
            malformed.body_table_header_settings(0usize),
            Err(Error::InvalidSource) | Err(Error::LimitExceeded { .. })
        ));
        assert_eq!(malformed.exact_bytes(), before.as_slice());
    }
    Ok(())
}

#[test]
fn locked_and_count_dependent_tables_refuse_partition_changes_but_allow_freeze() -> TestResult<()> {
    let locked_source = synthetic_package(["Revenue", "Costs"], Some(true))?;
    let locked = Package::from_bytes(&locked_source)?;
    let locked_error = locked
        .edit_body_table_header_settings(0usize)?
        .set(changed_settings())
        .commit()
        .expect_err("a locked table must reject a changed header edit");
    assert!(matches!(locked_error, Error::TableLocked));
    assert_eq!(locked.exact_bytes(), locked_source.as_slice());

    let dependency_source =
        append_dependency_field(&synthetic_package(["Revenue", "Costs"], None)?)?;
    let dependent = Package::from_bytes(&dependency_source)?;
    let result = dependent
        .edit_body_table_header_settings(0usize)
        .and_then(|edit| edit.set(changed_settings()).commit());
    assert!(result.is_err(), "dependent topology must fail closed");
    assert_eq!(dependent.exact_bytes(), dependency_source.as_slice());

    let before = dependent.body_table_header_settings(0usize)?;
    let freeze_only = Settings {
        header_rows_frozen: Some(true),
        ..before
    };
    let commit = dependent
        .edit_body_table_header_settings(0usize)?
        .set(freeze_only)
        .commit()?;
    assert_eq!(
        commit.package().body_table_header_settings(0usize)?,
        freeze_only
    );
    Ok(())
}

#[test]
fn malformed_dependency_scalars_and_external_references_fail_closed() -> TestResult<()> {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let mut malformed_model_dependency = Vec::new();
    let mut category_group = Vec::new();
    category_group.extend_from_slice(&[0x30, 0x81, 0x00]);
    let mut category_owner = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut category_owner,
        2,
        &category_group,
    )?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut malformed_model_dependency,
        81,
        &category_owner,
    )?;
    for malformed_source in [
        append_table_info_raw(&source, &[0x38, 0x01])?,
        append_table_info_raw(&source, &[0x80, 0x01, 0x81, 0x00])?,
        append_selected_model_raw(&source, &malformed_model_dependency)?,
        make_table_model_reference_external(&source)?,
    ] {
        let package = Package::from_bytes(&malformed_source)?;
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_header_settings(0usize)
            .and_then(|edit| {
                let settings = Settings {
                    header_rows_frozen: Some(true),
                    ..edit.settings()
                };
                edit.set(settings).commit()
            });
        assert!(matches!(result, Err(Error::InvalidSource)));
        assert_eq!(package.exact_bytes(), before.as_slice());
    }
    Ok(())
}

#[test]
fn finite_limits_fail_closed_before_publication() -> TestResult<()> {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let mut transaction_rejected = false;
    for fields in 1..=4_096 {
        let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        let Ok(package) = Package::from_bytes_with_limits(&source, limits) else {
            continue;
        };
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_header_settings(0usize)
            .and_then(|edit| edit.set(changed_settings()).commit());
        if matches!(result, Err(Error::LimitExceeded { .. })) {
            assert_eq!(package.exact_bytes(), before.as_slice());
            transaction_rejected = true;
            break;
        }
    }
    assert!(
        transaction_rejected,
        "a finite ingress-compatible budget must reject the edit"
    );
    Ok(())
}

#[test]
fn public_transaction_types_are_send_sync_debug_and_redacted() -> TestResult<()> {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Settings>();
    assert_send_sync_debug::<litchi_pages::BodyTableHeaderSettingsEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableHeaderSettingsPatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableHeaderSettingsCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableHeaderSettingsDiagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<litchi_pages::BodyTableHeaderSettingsLimitKind>();
    let package = Package::from_bytes(&synthetic_package(["Revenue", "Costs"], None)?)?;
    let edit = package.edit_body_table_header_settings(0usize)?;
    assert!(format!("{edit:?}").contains("settings"));
    Ok(())
}
