//! Exact-source integration coverage for persisted Numbers table sort rules.
//!
//! This suite intentionally stops at `TableModelArchive.sort_order` (field
//! 44).  It does not exercise the physical "Sort Now" operation, which moves
//! cells, tiles, UID maps, comments, and style sidecars and remains owned by
//! the compatibility host.  The fixture is the tracked native Numbers table
//! with a synthetic field-44 payload injected into the selected Table 1 model.

use std::{fmt::Debug, path::PathBuf};

use litchi_iwa_archive::{
    Limits as ArchiveLimits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsp, tst};
use litchi_numbers::{
    Package, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    table::{
        lock::State as LockState,
        sort::transaction::{Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path},
        sort::{ColumnIndex, Direction, Order, Rule, Scope},
    },
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const TABLE_MODEL_TYPE: u32 = 6_001;
const SORT_ORDER_FIELD: u32 = 44;
const SORT_TRACKER_FIELD: u32 = 45;
const SORT_TYPE_FIELD: u32 = 1;
const SORT_RULES_FIELD: u32 = 2;
const RULE_COLUMN_FIELD: u32 = 1;
const RULE_DIRECTION_FIELD: u32 = 2;
const CANONICAL_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

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

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn table_one_model(data: &[u8]) -> bool {
    tst::TableModelArchive::decode(data).is_ok_and(|model| model.table_name == "Table 1")
}

fn rewrite_native_table_model(
    source: &[u8],
    mutate: impl FnOnce(&mut ArchiveObject, usize, &mut Vec<u8>) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut selected = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        if archive.objects.iter().any(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == TABLE_MODEL_TYPE && table_one_model(&message.data))
        }) {
            selected = Some((entry.name().to_owned(), archive));
            break;
        }
    }
    let (member, mut archive) =
        selected.ok_or_else(|| std::io::Error::other("native Table 1 model is missing"))?;
    let (object_index, message_index) = archive
        .objects
        .iter()
        .enumerate()
        .find_map(|(object_index, object)| {
            object
                .messages
                .iter()
                .enumerate()
                .find_map(|(message_index, message)| {
                    (message.type_ == TABLE_MODEL_TYPE && table_one_model(&message.data))
                        .then_some((object_index, message_index))
                })
        })
        .ok_or_else(|| std::io::Error::other("native Table 1 payload is missing"))?;
    let object = &mut archive.objects[object_index];
    let mut model = object.messages[message_index].data.clone();
    mutate(object, message_index, &mut model)?;
    object.replace_message_preserving_header(
        message_index,
        RawMessage {
            type_: TABLE_MODEL_TYPE,
            data: model,
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&member, &compressed)],
        ArchiveLimits::default(),
    )?)
}

fn model_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in archive.objects {
            for message in object.messages {
                if message.type_ == TABLE_MODEL_TYPE && table_one_model(&message.data) {
                    return Ok(message.data);
                }
            }
        }
    }
    Err(std::io::Error::other("native Table 1 payload is missing").into())
}

fn without_sort(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_native_table_model(source, |_object, _message_index, model| {
        let data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        *model = rewrite_repeated_length_delimited_fields(&data, SORT_TRACKER_FIELD, &[])?;
        Ok(())
    })
}

fn sort_payload(scope: Scope, rules: &[Rule], with_unknowns: bool) -> TestResult<Vec<u8>> {
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
        let rules = repeated_length_delimited_payloads(&payload, SORT_RULES_FIELD)?;
        let first = rules
            .first()
            .ok_or_else(|| std::io::Error::other("sort rule is missing"))?;
        let mut first = first.to_vec();
        append_overlong_varint_field(&mut first, 92, 0);
        append_balanced_unknown_group(&mut first, 93);
        let mut rewritten = Vec::new();
        for (index, raw) in rules.into_iter().enumerate() {
            append_length_delimited_field(
                &mut rewritten,
                SORT_RULES_FIELD,
                if index == 0 { &first } else { raw },
            )?;
        }
        let mut fields = rewrite_repeated_length_delimited_fields(&payload, SORT_RULES_FIELD, &[])?;
        fields.extend_from_slice(&rewritten);
        append_overlong_varint_field(&mut fields, 90, 0);
        append_balanced_unknown_group(&mut fields, 91);
        payload = fields;
    }
    Ok(payload)
}

fn with_sort(
    source: &[u8],
    scope: Scope,
    rules: &[Rule],
    with_unknowns: bool,
    tracker: bool,
) -> TestResult<Vec<u8>> {
    let payload = sort_payload(scope, rules, with_unknowns)?;
    rewrite_native_table_model(source, |_object, _message_index, model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        data = rewrite_repeated_length_delimited_fields(&data, SORT_TRACKER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, &payload)?;
        if tracker {
            let tracker = tst::SortRuleReferenceTrackerArchive {
                reference_tracker: reference(91),
            }
            .encode_to_vec();
            append_length_delimited_field(&mut data, SORT_TRACKER_FIELD, &tracker)?;
        }
        *model = data;
        Ok(())
    })
}

fn with_duplicate_sort_field(source: &[u8]) -> TestResult<Vec<u8>> {
    let payload = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    rewrite_native_table_model(source, |_object, _message_index, model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, &payload)?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, &payload)?;
        *model = data;
        Ok(())
    })
}

fn with_wrong_wire_sort_field(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_native_table_model(source, |_object, _message_index, model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_varint_field(&mut data, SORT_ORDER_FIELD, 1)?;
        *model = data;
        Ok(())
    })
}

fn with_malformed_sort_payload(
    source: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>),
) -> TestResult<Vec<u8>> {
    let payload = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let mut payload = payload;
    mutate(&mut payload);
    rewrite_native_table_model(source, |_object, _message_index, model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, &payload)?;
        *model = data;
        Ok(())
    })
}

fn append_overlong_varint_field(data: &mut Vec<u8>, field: u32, value: u64) {
    push_varint(data, u64::from(field) << 3);
    if value == 0 {
        data.extend_from_slice(&[0x80, 0x00]);
        return;
    }
    push_varint(data, value);
}

fn append_balanced_unknown_group(data: &mut Vec<u8>, field: u32) {
    push_varint(data, u64::from(field) << 3 | 3);
    push_varint(data, 1 << 3);
    push_varint(data, 1);
    push_varint(data, u64::from(field) << 3 | 4);
}

fn push_varint(data: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        data.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    data.push(value as u8);
}

fn assert_model_field45_equal(source: &[u8], target: &[u8]) -> TestResult {
    let before = model_payload(source)?;
    let after = model_payload(target)?;
    assert_eq!(
        repeated_length_delimited_payloads(&before, SORT_TRACKER_FIELD)?,
        repeated_length_delimited_payloads(&after, SORT_TRACKER_FIELD)?,
        "sort_rule_reference_tracker field 45 changed"
    );
    Ok(())
}

fn assert_exact_locality(source: &[u8], target: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source)?;
    let target_catalog = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for before in source_catalog.iter() {
        let after = target_catalog
            .iter()
            .find(|candidate| candidate.name() == before.name())
            .ok_or_else(|| std::io::Error::other("candidate removed a source member"))?;
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
        if before.name().contains("Metadata") || CANONICAL_PREVIEWS.contains(&before.name()) {
            assert_eq!(
                before.data(),
                after.data(),
                "metadata or preview member {} changed",
                before.name()
            );
        }
    }
    assert_eq!(
        changed.len(),
        1,
        "sort configuration should rewrite one member"
    );
    assert!(changed[0].ends_with(".iwa"));
    assert_eq!(source_catalog.len(), target_catalog.len());
    Ok(())
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        // A strict owner may reject malformed field-44 input at ingress.
        return Ok(());
    };
    let before = package.exact_bytes();
    assert!(package.table_sort_order(0usize, 0usize).is_err());
    assert!(package.edit_table_sort_order(0usize, 0usize).is_err());
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn order(column: usize, direction: Direction) -> TestResult<Order> {
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

#[test]
fn transaction_types_are_typed_and_debuggable() {
    fn assert_send_sync_debug<T: Send + Sync + Debug>() {}
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<ColumnIndex>();
    assert_send_sync_debug::<Direction>();
    assert_send_sync_debug::<Order>();
    assert_send_sync_debug::<Rule>();
    assert_send_sync_debug::<Scope>();
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();
}

#[test]
fn absent_sort_reads_none_and_clear_is_exact_noop() -> TestResult {
    let source = without_sort(&std::fs::read(fixture_path())?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.table_sort_order(0usize, 0usize)?, None);
    let commit = package
        .edit_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before().cloned(), None);
    assert_eq!(commit.patch().after().cloned(), None);
    assert_eq!(commit.package().exact_bytes(), source);
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    Ok(())
}

#[test]
fn set_reads_selected_scope_and_preserves_field45_metadata_previews_and_locality() -> TestResult {
    let source = with_sort(
        &std::fs::read(fixture_path())?,
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        true,
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let before = package.table_sort_order(0usize, 0usize)?;
    assert_eq!(
        before,
        Some(order(1, Direction::Ascending)?),
        "the injected native marker should be readable"
    );
    let expected = two_rule_order()?;
    let commit = package
        .edit_table_sort_order(0usize, 0usize)?
        .set(expected.clone())
        .commit()?;
    assert_eq!(commit.patch().path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(commit.patch().before().cloned(), before);
    assert_eq!(commit.patch().after().cloned(), Some(expected.clone()));
    assert_eq!(
        commit.package().table_sort_order(0usize, 0usize)?,
        Some(expected)
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    let target = commit.package().exact_bytes();
    assert_ne!(target, source);
    assert_exact_locality(&source, &target)?;
    assert_model_field45_equal(&source, &target)?;
    let target_model = model_payload(&target)?;
    let sort = repeated_length_delimited_payloads(&target_model, SORT_ORDER_FIELD)?;
    let sort = sort
        .first()
        .ok_or_else(|| std::io::Error::other("sort payload disappeared"))?;
    assert!(sort.windows(2).any(|window| window == [0x80, 0x00]));
    assert!(sort.windows(2).any(|window| window == [0xdb, 0x05]));
    Ok(())
}

#[test]
fn empty_marker_reads_none_and_clear_preserves_scope_unknowns_tracker_and_locality() -> TestResult {
    let source = with_sort(
        &std::fs::read(fixture_path())?,
        Scope::SelectedRows,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Descending)],
        true,
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    assert_eq!(commit.package().table_sort_order(0usize, 0usize)?, None);
    let cleared_bytes = commit.package().exact_bytes();
    let model = model_payload(&cleared_bytes)?;
    let sort_payloads = repeated_length_delimited_payloads(&model, SORT_ORDER_FIELD)?;
    let sort = sort_payloads
        .first()
        .ok_or_else(|| std::io::Error::other("clear removed native empty marker"))?;
    let sort = (*sort).to_vec();
    let native = tst::TableSortOrderArchive::decode(sort.as_slice())?;
    assert_eq!(native.r#type, Scope::SelectedRows.native_value());
    assert!(native.rules.is_empty());
    assert!(sort.windows(2).any(|window| window == [0x80, 0x00]));
    let target = commit.package().exact_bytes();
    assert_model_field45_equal(&source, &target)?;
    assert_exact_locality(&source, &target)?;
    Ok(())
}

#[test]
fn set_patch_apply_inverse_and_conflict_are_exact_source_bound() -> TestResult {
    let source = without_sort(&std::fs::read(fixture_path())?)?;
    let package = Package::from_bytes(&source)?;
    let expected = order(1, Direction::Ascending)?;
    let commit = package
        .edit_table_sort_order(0usize, 0usize)?
        .set(expected.clone())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        package
            .apply_table_sort_order(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    assert!(matches!(
        reopened.apply_table_sort_order(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = reopened.apply_table_sort_order(&inverse)?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(restored.package().table_sort_order(0usize, 0usize)?, None);
    let noop = reopened
        .edit_table_sort_order(0usize, 0usize)?
        .set(expected)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), target);
    Ok(())
}

#[test]
fn selected_rows_scope_roundtrips_without_inventing_a_persisted_range() -> TestResult {
    let source = without_sort(&std::fs::read(fixture_path())?)?;
    let selected = Order::selected_rows([Rule::new(ColumnIndex::new(1)?, Direction::Descending)])?;
    let commit = Package::from_bytes(&source)?
        .edit_table_sort_order(0usize, 0usize)?
        .set(selected.clone())
        .commit()?;
    assert_eq!(
        commit.package().table_sort_order(0usize, 0usize)?,
        Some(selected.clone())
    );
    assert_eq!(commit.patch().before().cloned(), None);
    assert_eq!(commit.patch().after().cloned(), Some(selected));
    Ok(())
}

#[test]
fn changed_sort_rejects_locked_table_atomically() -> TestResult {
    let source = without_sort(&std::fs::read(fixture_path())?)?;
    let unlocked = Package::from_bytes(&source)?;
    let mut lock_edit = unlocked.edit_table_lock(0usize, 0usize)?;
    lock_edit.lock();
    let locked = lock_edit.commit()?.into_package();
    assert_eq!(locked.table_lock(0usize, 0usize)?, LockState::Locked);
    let before = locked.exact_bytes();
    let error = locked
        .edit_table_sort_order(0usize, 0usize)?
        .set(order(1, Direction::Ascending)?)
        .commit()
        .expect_err("changed persisted sort must reject a locked table");
    assert!(matches!(error, Error::TableLocked { .. }));
    assert_eq!(locked.exact_bytes(), before);
    Ok(())
}

#[test]
fn malformed_sort_wire_rejects_duplicate_and_wrong_wire_fields_atomically() -> TestResult {
    let fixture = std::fs::read(fixture_path())?;
    for source in [
        with_duplicate_sort_field(&fixture)?,
        with_wrong_wire_sort_field(&fixture)?,
    ] {
        assert_rejected_atomically(&source)?;
    }
    Ok(())
}

#[test]
fn malformed_sort_wire_rejects_duplicate_known_nested_fields_and_unknown_enums() -> TestResult {
    let fixture = std::fs::read(fixture_path())?;
    let duplicate_type = with_malformed_sort_payload(&fixture, |payload| {
        append_varint_field(
            payload,
            SORT_TYPE_FIELD,
            Scope::EntireTable.native_value() as u64,
        )
        .expect("append duplicate type");
    })?;
    let duplicate_rule_column = with_malformed_sort_payload(&fixture, |payload| {
        let rules =
            repeated_length_delimited_payloads(payload, SORT_RULES_FIELD).expect("read sort rule");
        let first = rules.first().expect("sort rule").to_vec();
        let mut malformed = first;
        append_varint_field(&mut malformed, RULE_COLUMN_FIELD, 1).expect("append duplicate column");
        let mut rewritten =
            rewrite_repeated_length_delimited_fields(payload, SORT_RULES_FIELD, &[])
                .expect("clear sort rule");
        append_length_delimited_field(&mut rewritten, SORT_RULES_FIELD, &malformed)
            .expect("append malformed rule");
        *payload = rewritten;
    })?;
    let duplicate_rule_direction = with_malformed_sort_payload(&fixture, |payload| {
        let rules =
            repeated_length_delimited_payloads(payload, SORT_RULES_FIELD).expect("read sort rule");
        let first = rules.first().expect("sort rule").to_vec();
        let mut malformed = first;
        append_varint_field(&mut malformed, RULE_DIRECTION_FIELD, 0)
            .expect("append duplicate direction");
        let mut rewritten =
            rewrite_repeated_length_delimited_fields(payload, SORT_RULES_FIELD, &[])
                .expect("clear sort rule");
        append_length_delimited_field(&mut rewritten, SORT_RULES_FIELD, &malformed)
            .expect("append malformed rule");
        *payload = rewritten;
    })?;
    let unknown_scope = with_malformed_sort_payload(&fixture, |payload| {
        // Replace the canonical type field with an unknown enum value.
        let rewritten =
            litchi_iwa_common::wire::patch_varint_field(payload, SORT_TYPE_FIELD, true, Some(99))
                .expect("patch unknown scope");
        *payload = rewritten;
    })?;
    for source in [
        duplicate_type,
        duplicate_rule_column,
        duplicate_rule_direction,
        unknown_scope,
    ] {
        assert_rejected_atomically(&source)?;
    }
    Ok(())
}

#[test]
fn unknown_overlong_scalar_and_balanced_group_bytes_survive_set_and_clear() -> TestResult {
    let fixture = std::fs::read(fixture_path())?;
    let source = with_sort(
        &fixture,
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        true,
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let changed = package
        .edit_table_sort_order(0usize, 0usize)?
        .set(two_rule_order()?)
        .commit()?;
    let changed_bytes = changed.package().exact_bytes();
    let changed_model = model_payload(&changed_bytes)?;
    let changed_sort = repeated_length_delimited_payloads(&changed_model, SORT_ORDER_FIELD)?
        .first()
        .ok_or_else(|| std::io::Error::other("changed sort payload missing"))?
        .to_vec();
    assert!(changed_sort.windows(2).any(|window| window == [0x80, 0x00]));
    assert!(changed_sort.windows(2).any(|window| window == [0xdb, 0x05]));
    let cleared = changed
        .package()
        .edit_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    let cleared_bytes = cleared.package().exact_bytes();
    let cleared_model = model_payload(&cleared_bytes)?;
    let cleared_sorts = repeated_length_delimited_payloads(&cleared_model, SORT_ORDER_FIELD)?;
    let cleared_sort = cleared_sorts
        .first()
        .ok_or_else(|| std::io::Error::other("cleared sort marker missing"))?;
    assert!(cleared_sort.windows(2).any(|window| window == [0x80, 0x00]));
    assert!(cleared_sort.windows(2).any(|window| window == [0xdb, 0x05]));
    Ok(())
}

#[test]
fn public_limits_reject_input_and_reference_ingress_before_publication() -> TestResult {
    let fixture = std::fs::read(fixture_path())?;
    let source = without_sort(&fixture)?;
    let too_small_input = PackageLimits::new(
        source.len().saturating_sub(1) as u64,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(too_small_input, PackageSemanticLimits::default()),
        )
        .is_err()
    );

    let too_few_references = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        PackageSemanticLimits::MAX_SHEETS,
        PackageSemanticLimits::MAX_TABLES,
        1,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(PackageLimits::default(), too_few_references),
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn selector_errors_are_redacted_and_public_values_are_archive_free() -> TestResult {
    let package = Package::from_bytes(&without_sort(&std::fs::read(fixture_path())?)?)?;
    assert!(matches!(
        package.table_sort_order("missing sheet", 0usize),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_sort_order(0usize, "missing table"),
        Err(Error::TableNotFound)
    ));
    let order = order(1, Direction::Ascending)?;
    let rendered = format!("{order:?}");
    assert!(!rendered.contains("Index/"));
    assert!(!rendered.contains("Document.iwa"));
    Ok(())
}
