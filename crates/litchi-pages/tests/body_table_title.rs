//! Native integration coverage for selector-first Pages body-table titles.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{decode_varint_from_bytes, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsd, tsp, tss, tst, tswp};
use litchi_pages::table::title::{Error, Settings};
use litchi_pages::{BodyTableSelector, Limits, Package};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const BODY_IDENTIFIER: u64 = 42;
const FIRST_ATTACHMENT_IDENTIFIER: u64 = 100;
const SECOND_ATTACHMENT_IDENTIFIER: u64 = 110;
const FIRST_DRAWABLE_IDENTIFIER: u64 = 200;
const SECOND_DRAWABLE_IDENTIFIER: u64 = 210;
const FIRST_MODEL_IDENTIFIER: u64 = 300;
const SECOND_MODEL_IDENTIFIER: u64 = 310;
const TITLE_STYLE_IDENTIFIER: u64 = 600;
const TITLE_SHAPE_STYLE_IDENTIFIER: u64 = 601;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn reference(identifier: u64) -> tsp::Reference {
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

fn object(
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

fn title_model(name: &str, visible: Option<bool>, outlined: Option<bool>) -> Vec<u8> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        table_name_enabled: visible,
        table_name_height: Some(20.0),
        table_name_border_enabled: outlined,
        table_name_style: Some(reference(TITLE_STYLE_IDENTIFIER)),
        table_name_shape_style: Some(reference(TITLE_SHAPE_STYLE_IDENTIFIER)),
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut payload,
        UNKNOWN_MODEL_FIELD,
        UNKNOWN_MODEL_VALUE,
    )
    .expect("synthetic unknown title field fits");
    payload
}

fn paragraph_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            style_identifier: Some("body-table-title".to_owned()),
            ..tss::StyleArchive::default()
        },
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn shape_style_payload() -> Vec<u8> {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss::StyleArchive {
                style_identifier: Some("body-table-title-shape".to_owned()),
                ..tss::StyleArchive::default()
            },
            ..tsd::ShapeStyleArchive::default()
        },
        ..tswp::ShapeStyleArchive::default()
    }
    .encode_to_vec()
}

fn table_ids(index: usize) -> (u64, u64, u64) {
    match index {
        0 => (
            FIRST_ATTACHMENT_IDENTIFIER,
            FIRST_DRAWABLE_IDENTIFIER,
            FIRST_MODEL_IDENTIFIER,
        ),
        1 => (
            SECOND_ATTACHMENT_IDENTIFIER,
            SECOND_DRAWABLE_IDENTIFIER,
            SECOND_MODEL_IDENTIFIER,
        ),
        _ => panic!("synthetic fixture only supports two tables"),
    }
}

fn synthetic_package(names: [&str; 2]) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let entries = names
        .iter()
        .enumerate()
        .map(|(index, _)| {
            let (attachment, _, _) = table_ids(index);
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
        let (attachment_identifier, drawable_identifier, model_identifier) = table_ids(index);
        let attachment_payload = tswp::DrawableAttachmentArchive {
            drawable: Some(reference(drawable_identifier)),
            ..tswp::DrawableAttachmentArchive::default()
        }
        .encode_to_vec();
        let mut table_info_payload = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(BODY_IDENTIFIER)),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(model_identifier),
            ..tst::TableInfoArchive::default()
        }
        .encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(&mut table_info_payload, 98, 17)?;

        let mut attachment_object = object(
            attachment_identifier,
            ATTACHMENT_MESSAGE_TYPE,
            attachment_payload,
            &[drawable_identifier],
        )?;
        attachment_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![1], drawable_identifier));

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

        let mut model_object = object(
            model_identifier,
            TABLE_MODEL_MESSAGE_TYPE,
            title_model(name, Some(true), Some(false)),
            &[TITLE_STYLE_IDENTIFIER, TITLE_SHAPE_STYLE_IDENTIFIER],
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![30], TITLE_STYLE_IDENTIFIER),
                field_reference(vec![36], TITLE_SHAPE_STYLE_IDENTIFIER),
            ]);
        objects.extend([attachment_object, drawable_object, model_object]);
    }
    objects.push(object(
        TITLE_STYLE_IDENTIFIER,
        2_022,
        paragraph_style_payload(),
        &[],
    )?);
    objects.push(object(
        TITLE_SHAPE_STYLE_IDENTIFIER,
        2_025,
        shape_style_payload(),
        &[],
    )?);
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

fn rewrite_document_archive(
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

fn append_duplicate_visible_field(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        litchi_iwa_common::wire::append_varint_field(&mut message.data, 22, 1)?;
        Ok(())
    })
}

fn remove_title_style_object(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(TITLE_STYLE_IDENTIFIER));
        Ok(())
    })
}

fn replace_title_style_message_type(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TITLE_STYLE_IDENTIFIER)
            .ok_or("missing title style object")?;
        object.messages[0].type_ = 2_023;
        object.archive_info.message_infos[0].type_ = 2_023;
        Ok(())
    })
}

fn remove_model_style_ownership(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model object")?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or("missing model message info")?;
        info.object_references.clear();
        info.field_infos.clear();
        Ok(())
    })
}

fn tamper_model_style_ownership(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model object")?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or("missing model message info")?;
        let reference = info
            .object_references
            .first_mut()
            .ok_or("missing model style reference")?;
        *reference = 999;
        Ok(())
    })
}

fn aggregate_only_source(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        for (identifier, type_) in [
            (1, ROOT_MESSAGE_TYPE),
            (BODY_IDENTIFIER, BODY_MESSAGE_TYPE),
            (FIRST_ATTACHMENT_IDENTIFIER, ATTACHMENT_MESSAGE_TYPE),
            (FIRST_DRAWABLE_IDENTIFIER, TABLE_INFO_MESSAGE_TYPE),
            (FIRST_MODEL_IDENTIFIER, TABLE_MODEL_MESSAGE_TYPE),
        ] {
            let object = archive.object_mut(identifier).ok_or("missing object")?;
            let index = object
                .messages
                .iter()
                .position(|message| message.type_ == type_)
                .ok_or("missing message")?;
            object.archive_info.message_infos[index].field_infos.clear();
        }
        let table_info = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table info")?;
        table_info.archive_info.message_infos[0]
            .object_references
            .retain(|identifier| *identifier != BODY_IDENTIFIER);
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

#[test]
fn selectors_read_position_and_name_and_reject_ambiguity() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let by_position = package.body_table_title_settings(BodyTableSelector::index(0))?;
    let by_name = package.body_table_title_settings(BodyTableSelector::name("Revenue"))?;
    assert_eq!(by_position, by_name);
    assert_eq!(by_position, Settings::new(Some(true), Some(false)));
    assert!(matches!(
        package.body_table_title_settings(BodyTableSelector::name("Missing")),
        Err(Error::TableNotFound)
    ));

    let ambiguous = Package::from_bytes(&synthetic_package(["Revenue", "Revenue"])?)?;
    assert!(matches!(
        ambiguous.body_table_title_settings(BodyTableSelector::name("Revenue")),
        Err(Error::AmbiguousTableName | Error::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn no_op_is_exact_and_presence_preserving() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_title_settings(0usize)?;
    let commit = package
        .edit_body_table_title(BodyTableSelector::name("Revenue"))?
        .set(before)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(commit.package().source_bytes(), source.as_slice());
    assert_eq!(
        commit
            .package()
            .apply_body_table_title(commit.patch())?
            .package()
            .source_bytes(),
        source.as_slice()
    );
    Ok(())
}

#[test]
fn changed_presence_preserves_unknowns_deletes_previews_and_inverts_exactly() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let after = Settings::new(Some(false), None);
    let commit = package.edit_body_table_title(0usize)?.set(after).commit()?;
    assert_eq!(
        commit.package().body_table_title_settings("Revenue")?,
        after
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        sentinel(commit.package().source_bytes())?,
        b"untouched-sentinel"
    );
    let changed_model = model_payload(commit.package().source_bytes(), FIRST_MODEL_IDENTIFIER)?;
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
        .apply_body_table_title(&commit.patch().inverse())?;
    assert_eq!(restored.package().source_bytes(), source.as_slice());
    assert_eq!(
        restored.package().body_table_title_settings(0usize)?,
        Settings::new(Some(true), Some(false))
    );
    Ok(())
}

#[test]
fn patch_conflict_and_malformed_selected_fields_are_atomic() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"])?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_title(0usize)?
        .set(Settings::new(Some(false), None))
        .commit()?;
    let catalog = Catalog::from_bytes(&source)?;
    let tampered_source = catalog.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_title(commit.patch()),
        Err(Error::PatchConflict)
    ));

    let malformed_source = append_duplicate_visible_field(&source)?;
    let malformed = Package::from_bytes(&malformed_source)?;
    assert!(matches!(
        malformed.body_table_title_settings(0usize),
        Err(Error::InvalidSource | Error::LimitExceeded { .. })
    ));
    assert!(matches!(
        malformed
            .edit_body_table_title(0usize)
            .and_then(|edit| edit.set(Settings::new(Some(false), None)).commit()),
        Err(Error::InvalidSource | Error::LimitExceeded { .. })
    ));
    assert_eq!(malformed.source_bytes(), malformed_source.as_slice());
    Ok(())
}

#[test]
fn title_dependencies_and_model_ownership_fail_closed_atomically() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"])?;
    for (label, malformed) in [
        ("missing style object", remove_title_style_object(&source)?),
        (
            "wrong style message type",
            replace_title_style_message_type(&source)?,
        ),
        (
            "missing model style ownership",
            remove_model_style_ownership(&source)?,
        ),
        (
            "tampered model style ownership",
            tamper_model_style_ownership(&source)?,
        ),
    ] {
        let package = Package::from_bytes(&malformed)?;
        let result = package
            .edit_body_table_title(0usize)
            .and_then(|edit| edit.set(Settings::new(Some(true), Some(true))).commit());
        assert!(
            matches!(result, Err(Error::InvalidSource)),
            "{label} unexpectedly published: {result:?}"
        );
        assert_eq!(package.source_bytes(), malformed.as_slice());
    }
    Ok(())
}

#[test]
fn aggregate_only_native_ownership_is_accepted_and_limits_fail_closed() -> TestResult {
    let source = aggregate_only_source(&synthetic_package(["Revenue", "Costs"])?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_title_settings(0usize)?,
        Settings::new(Some(true), Some(false))
    );
    let commit = package
        .edit_body_table_title(0usize)?
        .set(Settings::new(Some(false), None))
        .commit()?;
    assert_eq!(
        commit.package().body_table_title_settings(0usize)?,
        Settings::new(Some(false), None)
    );

    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    let bounded = Package::from_bytes_with_limits(&source, limits);
    if let Ok(bounded) = bounded {
        assert!(matches!(
            bounded.body_table_title_settings(0usize),
            Err(Error::LimitExceeded { .. }) | Err(Error::InvalidSource)
        ));
    }
    Ok(())
}

#[test]
fn public_transaction_values_are_send_sync_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Settings>();
    assert_send_sync_debug::<litchi_pages::BodyTableTitleEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableTitlePatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableTitleCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableTitleDiagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<litchi_pages::BodyTableTitleLimitKind>();
    let package = Package::from_bytes(&synthetic_package(["Revenue", "Costs"])?)?;
    let edit = package.edit_body_table_title(0usize)?;
    assert!(format!("{edit:?}").contains("settings"));
    Ok(())
}
