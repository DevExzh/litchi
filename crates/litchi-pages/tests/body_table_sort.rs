//! Exact-source integration coverage for persisted Pages body-table sort rules.
//!
//! The fixture is deliberately small and synthetic: it contains a rooted body
//! table, a persisted `TableModelArchive.sort_order` field 44, and the native
//! field-45 sort-rule tracker.  This suite covers only the configuration marker
//! and never performs the physical row-reordering operation.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsd, tsp, tst, tswp};
use litchi_pages::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};
use litchi_pages::{BodyTableSelector, Limits, Package};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const ROOT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const FIRST_ATTACHMENT_IDENTIFIER: u64 = 100;
const SECOND_ATTACHMENT_IDENTIFIER: u64 = 110;
const FIRST_DRAWABLE_IDENTIFIER: u64 = 200;
const SECOND_DRAWABLE_IDENTIFIER: u64 = 210;
const FIRST_MODEL_IDENTIFIER: u64 = 300;
const SECOND_MODEL_IDENTIFIER: u64 = 310;
const TRACKER_IDENTIFIER: u64 = 900;
const TRACKER_TARGET_IDENTIFIER: u64 = 901;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TRACKER_MESSAGE_TYPE: u32 = 6_010;
const SORT_ORDER_FIELD: u32 = 44;
const SORT_TRACKER_FIELD: u32 = 45;
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

fn sort_order(scope: Scope, rules: &[Rule], with_unknowns: bool) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableSortOrderArchive {
        r#type: scope.native_value(),
        rules: rules
            .iter()
            .map(|rule| tst::table_sort_order_archive::SortRuleArchive {
                index: rule.column().native_value(),
                direction: rule.direction().native_value(),
            })
            .collect(),
    }
    .encode_to_vec();

    if with_unknowns {
        append_overlong_varint_field(&mut payload, 90, 0);
        let rules = repeated_length_delimited_payloads(&payload, 2)?;
        if let Some(first) = rules.first() {
            let mut first = first.to_vec();
            append_overlong_varint_field(&mut first, 92, 0);
            let mut rewritten = rewrite_repeated_length_delimited_fields(&payload, 2, &[])?;
            for (index, rule) in rules.into_iter().enumerate() {
                append_length_delimited_field(
                    &mut rewritten,
                    2,
                    if index == 0 { first.as_slice() } else { rule },
                )?;
            }
            payload = rewritten;
        }
    }
    Ok(payload)
}

fn model_payload(
    name: &str,
    sort: Option<&[u8]>,
    with_unknowns: bool,
    with_tracker: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 3,
        number_of_columns: 3,
        base_data_store: tst::DataStore::default(),
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    if with_unknowns {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    if let Some(sort) = sort {
        append_length_delimited_field(&mut payload, SORT_ORDER_FIELD, sort)?;
    }
    if with_tracker {
        let tracker = tst::SortRuleReferenceTrackerArchive {
            reference_tracker: reference(TRACKER_TARGET_IDENTIFIER),
        }
        .encode_to_vec();
        append_length_delimited_field(&mut payload, SORT_TRACKER_FIELD, &tracker)?;
    }
    Ok(payload)
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

fn synthetic_package(names: [&str; 2], sort: Option<&[u8]>) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["\u{fffc}\u{fffc}".to_owned()],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: names
                .iter()
                .enumerate()
                .map(|(index, _)| {
                    let (attachment, _, _) = table_ids(index);
                    tswp::object_attribute_table::ObjectAttribute {
                        character_index: u32::try_from(index).expect("fixture index fits"),
                        object: Some(reference(attachment)),
                    }
                })
                .collect(),
        }),
        ..tswp::StorageArchive::default()
    };

    let mut root_object = object(
        ROOT_IDENTIFIER,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));

    let attachment_ids = [FIRST_ATTACHMENT_IDENTIFIER, SECOND_ATTACHMENT_IDENTIFIER];
    let mut body_object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body.encode_to_vec(),
        &attachment_ids,
    )?;
    body_object.archive_info.message_infos[0]
        .field_infos
        .extend(attachment_ids.map(|identifier| field_reference(vec![9], identifier)));

    let mut objects = vec![root_object, body_object];
    for (index, name) in names.iter().enumerate() {
        let (attachment_identifier, drawable_identifier, model_identifier) = table_ids(index);
        let attachment = tswp::DrawableAttachmentArchive {
            drawable: Some(reference(drawable_identifier)),
            ..tswp::DrawableAttachmentArchive::default()
        };
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(BODY_IDENTIFIER)),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(model_identifier),
            ..tst::TableInfoArchive::default()
        };
        let mut attachment_object = object(
            attachment_identifier,
            ATTACHMENT_MESSAGE_TYPE,
            attachment.encode_to_vec(),
            &[drawable_identifier],
        )?;
        attachment_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![1], drawable_identifier));

        let mut drawable_object = object(
            drawable_identifier,
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
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
            model_payload(name, sort, true, true)?,
            &[TRACKER_IDENTIFIER, TRACKER_TARGET_IDENTIFIER],
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![45, 1], TRACKER_TARGET_IDENTIFIER),
                field_reference(vec![99], TRACKER_IDENTIFIER),
            ]);
        objects.extend([attachment_object, drawable_object, model_object]);
    }
    objects.push(object(
        TRACKER_IDENTIFIER,
        TRACKER_MESSAGE_TYPE,
        tst::SortRuleReferenceTrackerArchive {
            reference_tracker: reference(TRACKER_TARGET_IDENTIFIER),
        }
        .encode_to_vec(),
        &[TRACKER_TARGET_IDENTIFIER],
    )?);
    objects.push(object(
        TRACKER_TARGET_IDENTIFIER,
        TRACKER_MESSAGE_TYPE + 1,
        vec![0x08, 0x01],
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

fn source_without_sort() -> TestResult<Vec<u8>> {
    synthetic_package(["Revenue", "Costs"], None)
}

fn source_with_sort(with_unknowns: bool) -> TestResult<Vec<u8>> {
    let sort = sort_order(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        with_unknowns,
    )?;
    synthetic_package(["Revenue", "Costs"], Some(&sort))
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

fn model_payload_from(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = document_archive(package)?;
    Ok(archive
        .object(identifier)
        .ok_or("missing table model")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table model message")?
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

fn rewrite_model(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut data = message.data.clone();
        mutate(&mut data)?;
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| format!("missing package member {name}"))?
        .data()
        .to_vec())
}

fn append_overlong_varint_field(data: &mut Vec<u8>, field: u32, value: u64) {
    push_varint(data, u64::from(field) << 3);
    if value == 0 {
        data.extend_from_slice(&[0x80, 0x00]);
    } else {
        push_varint(data, value);
    }
}

fn push_varint(data: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        data.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    data.push(value as u8);
}

fn append_duplicate_sort_field(package: &[u8]) -> TestResult<Vec<u8>> {
    let sort = sort_order(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    rewrite_model(package, |model| {
        append_length_delimited_field(model, SORT_ORDER_FIELD, &sort)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &sort)?;
        Ok(())
    })
}

fn append_wrong_wire_sort_field(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model(package, |model| {
        append_varint_field(model, SORT_ORDER_FIELD, 1)?;
        Ok(())
    })
}

fn append_noncanonical_nested_field(package: &[u8]) -> TestResult<Vec<u8>> {
    let sort = sort_order(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let mut malformed = sort;
    append_overlong_varint_field(&mut malformed, 1, 0);
    rewrite_model(package, |model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, &malformed)?;
        *model = data;
        Ok(())
    })
}

fn tracker_reference_payload(
    identifier: u64,
    reference_type: Option<u64>,
    external: Option<u64>,
    duplicate_identifier: bool,
) -> TestResult<Vec<u8>> {
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, identifier)?;
    if duplicate_identifier {
        append_varint_field(&mut reference, 1, identifier)?;
    }
    if let Some(reference_type) = reference_type {
        append_varint_field(&mut reference, 2, reference_type)?;
    }
    if let Some(external) = external {
        append_varint_field(&mut reference, 3, external)?;
    }
    Ok(reference)
}

fn tracker_payload(reference: &[u8]) -> TestResult<Vec<u8>> {
    let mut tracker = Vec::new();
    append_length_delimited_field(&mut tracker, 1, reference)?;
    Ok(tracker)
}

fn replace_tracker_payload(package: &[u8], tracker: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model(package, |model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_TRACKER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_TRACKER_FIELD, tracker)?;
        *model = data;
        Ok(())
    })
}

fn wrong_model_message_type(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table model message")?;
        message.type_ = TABLE_INFO_MESSAGE_TYPE;
        Ok(())
    })
}

fn duplicate_model_identifier(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(SECOND_MODEL_IDENTIFIER)
            .ok_or("missing second table model")?;
        model.archive_info.identifier = Some(FIRST_MODEL_IDENTIFIER);
        Ok(())
    })
}

fn cross_component_model_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let inbound = object(700, 7_000, vec![0x08, 0x01], &[FIRST_MODEL_IDENTIFIER])?;
    let component = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members = Vec::new();
    for entry in catalog.iter() {
        members.push((entry.name().to_owned(), entry.data().to_vec()));
    }
    members.push(("Index/Other.iwa".to_owned(), component));
    let references = members
        .iter()
        .map(|(name, data)| (name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        references,
        Limits::default(),
    )?)
}

fn aggregate_only_reference_metadata(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .field_infos
            .retain(|field| field.path.as_slice() != [SORT_TRACKER_FIELD, 1]);
        let table_info = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table info")?;
        table_info.archive_info.message_infos[0]
            .field_infos
            .retain(|field| field.path.as_slice() != [2]);
        Ok(())
    })
}

fn locked_source(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let info = archive
            .object_mut(FIRST_DRAWABLE_IDENTIFIER)
            .ok_or("missing table info")?;
        let message = info
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table info message")?;
        let mut native = tst::TableInfoArchive::decode(message.data.as_slice())?;
        native.super_.locked = Some(true);
        message.data = native.encode_to_vec();
        Ok(())
    })
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = package.exact_bytes();
    if package.body_table_sort_order(0usize).is_ok() {
        return Err("hostile package was accepted by the read route".into());
    }
    if package.edit_body_table_sort_order(0usize).is_ok() {
        return Err("hostile package was accepted by the edit route".into());
    }
    if package.exact_bytes() != before {
        return Err("hostile package changed during rejected ingress".into());
    }
    Ok(())
}

fn one_rule_order(column: usize, direction: Direction) -> TestResult<Order> {
    Ok(Order::new([Rule::new(
        ColumnIndex::new(column)?,
        direction,
    )])?)
}

fn two_rule_order() -> TestResult<Order> {
    Ok(Order::new([
        Rule::new(ColumnIndex::new(2)?, Direction::Descending),
        Rule::new(ColumnIndex::new(1)?, Direction::Ascending),
    ])?)
}

fn assert_field45_preserved(source: &[u8], target: &[u8]) -> TestResult {
    assert_eq!(
        repeated_length_delimited_payloads(
            &model_payload_from(source, FIRST_MODEL_IDENTIFIER)?,
            SORT_TRACKER_FIELD,
        )?,
        repeated_length_delimited_payloads(
            &model_payload_from(target, FIRST_MODEL_IDENTIFIER)?,
            SORT_TRACKER_FIELD,
        )?
    );
    Ok(())
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for before in source_catalog.iter() {
        let after = target_catalog
            .iter()
            .find(|candidate| candidate.name() == before.name())
            .ok_or("candidate removed source member")?;
        if before.data() != after.data() {
            changed.push(before.name().to_owned());
        } else {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record()
            );
            assert_eq!(before.data(), after.data(), "unselected member changed");
        }
        assert_eq!(before.name(), after.name());
    }
    assert_eq!(changed, vec![DOCUMENT_MEMBER.to_owned()]);
    assert_eq!(source_catalog.len(), target_catalog.len());
    Ok(())
}

#[test]
fn transaction_values_are_typed_and_redacted() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<ColumnIndex>();
    assert_send_sync_debug::<Direction>();
    assert_send_sync_debug::<Order>();
    assert_send_sync_debug::<Rule>();
    assert_send_sync_debug::<Scope>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortPatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortDiagnostics>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortError>();
    assert_send_sync_debug::<litchi_pages::BodyTableSortLimitKind>();
}

#[test]
fn absent_sort_reads_none_and_clear_is_exact_noop() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.body_table_sort_order(0usize)?, None);
    let commit = package
        .edit_body_table_sort_order(BodyTableSelector::name("Revenue"))?
        .clear()
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().exact_bytes(), source);
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    Ok(())
}

#[test]
fn selectors_read_by_index_and_name_and_reject_missing_or_ambiguous() -> TestResult {
    let source = synthetic_package(["Revenue", "Costs"], None)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.body_table_sort_order(BodyTableSelector::index(0))?,
        package.body_table_sort_order(BodyTableSelector::name("Revenue"))?
    );
    assert!(
        package
            .body_table_sort_order(BodyTableSelector::name("Missing"))
            .is_err()
    );
    let ambiguous = Package::from_bytes(&synthetic_package(["Revenue", "Revenue"], None)?)?;
    assert!(
        ambiguous
            .body_table_sort_order(BodyTableSelector::name("Revenue"))
            .is_err()
    );
    Ok(())
}

#[test]
fn set_replaces_sort_order_and_preserves_field45_unknowns_locality_and_previews() -> TestResult {
    let source = source_with_sort(true)?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_sort_order(0usize)?;
    assert_eq!(before, Some(one_rule_order(1, Direction::Ascending)?));
    let after = two_rule_order()?;
    let commit = package
        .edit_body_table_sort_order(BodyTableSelector::name("Revenue"))?
        .set(after.clone())
        .commit()?;
    assert_eq!(commit.package().body_table_sort_order(0usize)?, Some(after));
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    let target = commit.package().exact_bytes();
    assert_ne!(target, source);
    assert_locality(&source, &target)?;
    assert_field45_preserved(&source, &target)?;
    assert_eq!(
        member_bytes(&source, PREVIEWS[0])?,
        member_bytes(&target, PREVIEWS[0])?
    );
    let model = model_payload_from(&target, FIRST_MODEL_IDENTIFIER)?;
    assert!(model.windows(2).any(|window| window == [0x80, 0x00]));
    Ok(())
}

#[test]
fn set_apply_inverse_and_conflict_are_exact_and_atomic() -> TestResult {
    let source = source_without_sort()?;
    let expected = one_rule_order(1, Direction::Ascending)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_sort_order(0usize)?
        .set(expected.clone())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        package
            .apply_body_table_sort_order(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    assert!(
        reopened
            .apply_body_table_sort_order(commit.patch())
            .is_err()
    );
    let restored = reopened.apply_body_table_sort_order(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(restored.package().body_table_sort_order(0usize)?, None);
    Ok(())
}

#[test]
fn clear_removes_semantic_marker_but_preserves_tracker_and_inverse() -> TestResult {
    let source = source_with_sort(true)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_sort_order(0usize)?
        .clear()
        .commit()?;
    assert_eq!(commit.package().body_table_sort_order(0usize)?, None);
    assert_field45_preserved(&source, &commit.package().exact_bytes())?;
    let restored = commit
        .package()
        .apply_body_table_sort_order(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn selected_row_scope_roundtrips_without_a_persisted_range() -> TestResult {
    let source = source_without_sort()?;
    let selected = Order::selected_rows([Rule::new(ColumnIndex::new(1)?, Direction::Descending)])?;
    let commit = Package::from_bytes(&source)?
        .edit_body_table_sort_order(0usize)?
        .set(selected.clone())
        .commit()?;
    assert_eq!(
        commit.package().body_table_sort_order(0usize)?,
        Some(selected)
    );
    Ok(())
}

#[test]
fn locked_table_refuses_changed_sort_atomically() -> TestResult {
    let locked = locked_source(&source_without_sort()?)?;
    let package = Package::from_bytes(&locked)?;
    let before = package.exact_bytes();
    let result = package
        .edit_body_table_sort_order(0usize)?
        .set(one_rule_order(1, Direction::Ascending)?)
        .commit();
    assert!(result.is_err());
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_and_noncanonical_nested_fields_are_atomic() -> TestResult {
    let source = source_without_sort()?;
    for (case, malformed) in [
        (
            "duplicate sort field",
            append_duplicate_sort_field(&source)?,
        ),
        (
            "wrong-wire sort field",
            append_wrong_wire_sort_field(&source)?,
        ),
        (
            "noncanonical nested field",
            append_noncanonical_nested_field(&source)?,
        ),
    ] {
        assert_rejected_atomically(&malformed).map_err(|error| format!("{case}: {error}"))?;
    }
    Ok(())
}

#[test]
fn wrong_model_type_is_rejected_without_publication() -> TestResult {
    assert_rejected_atomically(&wrong_model_message_type(&source_without_sort()?)?)
}

#[test]
fn model_identifier_alias_and_cross_component_inbound_are_rejected() -> TestResult {
    let source = source_without_sort()?;
    if let Ok(duplicate) = duplicate_model_identifier(&source) {
        assert_rejected_atomically(&duplicate)
            .map_err(|error| format!("duplicate model identifier: {error}"))?;
    }
    let cross_component = cross_component_model_inbound(&source)?;
    assert_rejected_atomically(&cross_component)
        .map_err(|error| format!("cross-component model inbound: {error}"))?;
    Ok(())
}

#[test]
fn native_aggregate_only_reference_metadata_is_supported() -> TestResult {
    let source = aggregate_only_reference_metadata(&source_without_sort()?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.body_table_sort_order(0usize)?, None);
    let order = one_rule_order(1, Direction::Ascending)?;
    let commit = package
        .edit_body_table_sort_order(0usize)?
        .set(order.clone())
        .commit()?;
    assert_eq!(commit.package().body_table_sort_order(0usize)?, Some(order));
    Ok(())
}

#[test]
fn malformed_field45_reference_graph_is_rejected_atomically() -> TestResult {
    let source = source_without_sort()?;
    let valid_reference = tracker_reference_payload(TRACKER_TARGET_IDENTIFIER, None, None, false)?;
    let duplicate_outer = rewrite_model(&source, |model| {
        let tracker = tracker_payload(&valid_reference)?;
        append_length_delimited_field(model, SORT_TRACKER_FIELD, &tracker)?;
        Ok(())
    })?;
    let wrong_wire = rewrite_model(&source, |model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_TRACKER_FIELD, &[])?;
        append_varint_field(&mut data, SORT_TRACKER_FIELD, 1)?;
        *model = data;
        Ok(())
    })?;
    let duplicate_nested = replace_tracker_payload(
        &source,
        &tracker_payload(&tracker_reference_payload(
            TRACKER_TARGET_IDENTIFIER,
            None,
            None,
            true,
        )?)?,
    )?;
    let external = replace_tracker_payload(
        &source,
        &tracker_payload(&tracker_reference_payload(
            TRACKER_TARGET_IDENTIFIER,
            None,
            Some(1),
            false,
        )?)?,
    )?;
    let nonzero_type = replace_tracker_payload(
        &source,
        &tracker_payload(&tracker_reference_payload(
            TRACKER_TARGET_IDENTIFIER,
            Some(1),
            None,
            false,
        )?)?,
    )?;
    let missing_target = replace_tracker_payload(
        &source,
        &tracker_payload(&tracker_reference_payload(999_999, None, None, false)?)?,
    )?;
    for (name, malformed) in [
        ("duplicate outer field", duplicate_outer),
        ("wrong outer wire", wrong_wire),
        ("duplicate reference identifier", duplicate_nested),
        ("external reference", external),
        ("nonzero reference type", nonzero_type),
        ("missing target", missing_target),
    ] {
        assert_rejected_atomically(&malformed).map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

#[test]
fn field45_archive_info_mismatches_are_rejected_atomically() -> TestResult {
    let source = source_without_sort()?;
    let duplicate_aggregate = rewrite_document_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .object_references
            .push(TRACKER_TARGET_IDENTIFIER);
        Ok(())
    })?;
    let data_reference = rewrite_document_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .data_references
            .push(TRACKER_TARGET_IDENTIFIER);
        Ok(())
    })?;
    let wrong_field_path = rewrite_document_archive(&source, |archive| {
        let model = archive
            .object_mut(FIRST_MODEL_IDENTIFIER)
            .ok_or("missing table model")?;
        model.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![45], TRACKER_TARGET_IDENTIFIER));
        Ok(())
    })?;
    for (name, malformed) in [
        ("duplicate aggregate reference", duplicate_aggregate),
        ("data reference", data_reference),
        ("wrong field path", wrong_field_path),
    ] {
        assert_rejected_atomically(&malformed).map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

#[test]
fn unknown_overlong_scalars_survive_set_and_clear() -> TestResult {
    let source = source_with_sort(true)?;
    let package = Package::from_bytes(&source)?;
    let changed = package
        .edit_body_table_sort_order(0usize)?
        .set(two_rule_order()?)
        .commit()?;
    let model = model_payload_from(&changed.package().exact_bytes(), FIRST_MODEL_IDENTIFIER)?;
    assert!(model.windows(2).any(|window| window == [0x80, 0x00]));
    let cleared = changed
        .package()
        .edit_body_table_sort_order(0usize)?
        .clear()
        .commit()?;
    assert_field45_preserved(&source, &cleared.package().exact_bytes())?;
    Ok(())
}

#[test]
fn public_input_and_archive_field_limits_fail_before_publication() -> TestResult {
    let source = source_without_sort()?;
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let before = package.exact_bytes();
        assert!(package.body_table_sort_order(0usize).is_err());
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}
