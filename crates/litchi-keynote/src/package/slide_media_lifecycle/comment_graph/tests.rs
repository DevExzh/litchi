use std::path::PathBuf;

use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{FieldInfo, MessageInfo};

use super::*;

const COMMENT_COMPONENT: &str = "Index/Slide-2652150.iwa";
const NATIVE_ROOT: u64 = 2_653_723;
const NATIVE_COPY: u64 = 2_653_826;
const NATIVE_AUTHOR: u64 = 2_653_721;

fn native_package() -> Package {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/media-comments-duplicate-native.key"),
    )
    .expect("native comment fixture");
    Package::from_bytes(&bytes).expect("native comment package")
}

fn native_budget() -> LifecycleBudget {
    let package = native_package();
    LifecycleBudget::for_package(&package).expect("lifecycle budget")
}

fn native_limits(package: &Package) -> WireLimits {
    package
        .semantic_wire_limits()
        .expect("semantic wire limits")
}

#[test]
fn native_comment_plan_keeps_authors_as_dependencies() {
    let package = native_package();
    let limits = native_limits(&package);
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let plan = plan_comment_graph(
        &package,
        COMMENT_COMPONENT,
        NATIVE_ROOT,
        limits,
        &mut budget,
    )
    .expect("native comment plan");

    assert_eq!(plan.storage_ids, [NATIVE_ROOT]);
    assert_eq!(plan.author_ids, [NATIVE_AUTHOR]);
    assert_eq!(plan.author_dependencies.len(), 1);
    assert_eq!(plan.author_dependencies[0].storage_identifier, NATIVE_ROOT);
    assert_eq!(plan.author_dependencies[0].author_identifier, NATIVE_AUTHOR);
    assert_eq!(
        plan.author_dependencies[0].component_name.as_ref(),
        "Index/AnnotationAuthorStorage.iwa"
    );
    assert!(plan.root_storage_uuid().is_some());
}

#[test]
fn native_copied_comment_can_share_source_uuid() {
    let first = {
        let package = native_package();
        let limits = native_limits(&package);
        let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
        plan_comment_graph(
            &package,
            COMMENT_COMPONENT,
            NATIVE_ROOT,
            limits,
            &mut budget,
        )
        .expect("source comment plan")
    };
    let second = {
        let package = native_package();
        let limits = native_limits(&package);
        let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
        plan_comment_graph(
            &package,
            COMMENT_COMPONENT,
            NATIVE_COPY,
            limits,
            &mut budget,
        )
        .expect("native copy comment plan")
    };

    assert_eq!(first.root_storage_uuid(), second.root_storage_uuid());
    assert_eq!(first.author_ids, second.author_ids);
}

#[test]
fn unknown_root_payload_fields_reject_selected_comment_clone() {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, COMMENT_TEXT_FIELD, b"comment")
        .expect("text field");
    append_varint_field(&mut payload, 10_000, 1).expect("future scalar field");
    let package = native_package();
    let limits = native_limits(&package);
    let mut budget = native_budget();

    assert!(validate_comment_payload_wire(&payload, limits, &mut budget, true).is_err());
}

#[test]
fn unknown_reference_fields_fail_closed() {
    let mut reference = Vec::new();
    append_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, NATIVE_AUTHOR)
        .expect("reference identifier");
    append_varint_field(&mut reference, 99, 1).expect("unknown reference field");
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, COMMENT_AUTHOR_FIELD, &reference)
        .expect("author field");
    let package = native_package();
    let limits = native_limits(&package);
    let mut budget = native_budget();

    assert!(validate_comment_payload_wire(&payload, limits, &mut budget, true).is_err());
}

#[test]
fn unknown_date_and_uuid_fields_reject_selected_comment_clone() {
    let package = native_package();
    let limits = native_limits(&package);

    let mut date = Vec::new();
    // Fixed-width 64-bit fields have a one-byte key for these test numbers.
    date.push(((99_u32 << 3) | 1) as u8);
    date.extend_from_slice(&1_u64.to_le_bytes());
    let mut date_payload = Vec::new();
    append_length_delimited_field(&mut date_payload, COMMENT_DATE_FIELD, &date)
        .expect("date field");
    let mut budget = native_budget();
    assert!(validate_comment_payload_wire(&date_payload, limits, &mut budget, true).is_err());

    let mut uuid = Vec::new();
    append_varint_field(&mut uuid, 99, 1).expect("unknown UUID field");
    let mut uuid_payload = Vec::new();
    append_length_delimited_field(&mut uuid_payload, COMMENT_UUID_FIELD, &uuid)
        .expect("UUID field");
    let mut budget = native_budget();
    assert!(validate_comment_payload_wire(&uuid_payload, limits, &mut budget, true).is_err());
}

#[test]
fn direct_movie_comment_reference_rejects_legacy_and_unknown_fields() {
    let package = native_package();
    let limits = native_limits(&package);

    let mut canonical = Vec::new();
    append_varint_field(&mut canonical, REFERENCE_IDENTIFIER_FIELD, NATIVE_ROOT)
        .expect("reference identifier");
    let mut budget = native_budget();
    assert_eq!(
        super::super::graph::strict_reference_identifier(&canonical, limits, &mut budget, 3)
            .expect("canonical direct comment reference"),
        NATIVE_ROOT
    );

    let mut deprecated = canonical.clone();
    append_varint_field(&mut deprecated, REFERENCE_DEPRECATED_TYPE_FIELD, 1)
        .expect("deprecated reference field");
    let mut budget = native_budget();
    assert!(
        super::super::graph::strict_reference_identifier(&deprecated, limits, &mut budget, 3)
            .is_err()
    );

    let mut unknown = canonical;
    append_varint_field(&mut unknown, 99, 1).expect("unknown reference field");
    let mut budget = native_budget();
    assert!(
        super::super::graph::strict_reference_identifier(&unknown, limits, &mut budget, 3).is_err()
    );
}

#[test]
fn metadata_requires_exact_decoded_reference_aggregate() {
    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let facts = StorageFacts {
        author_identifier: Some(NATIVE_AUTHOR),
        reply_identifiers: vec![7, 9],
        uuid: None,
    };
    let mut info = MessageInfo::new(COMMENT_STORAGE_MESSAGE_TYPE, 1);
    info.object_references = vec![NATIVE_AUTHOR, 7, 9];
    validate_storage_metadata(1, &info, &facts, &mut budget).expect("complete aggregate");

    info.object_references.push(11);
    assert!(validate_storage_metadata(1, &info, &facts, &mut budget).is_err());
}

#[test]
fn field_metadata_must_be_a_multiset_subset_of_the_aggregate() {
    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let facts = StorageFacts {
        author_identifier: Some(NATIVE_AUTHOR),
        reply_identifiers: vec![7],
        uuid: None,
    };
    let mut info = MessageInfo::new(COMMENT_STORAGE_MESSAGE_TYPE, 1);
    info.object_references = vec![NATIVE_AUTHOR, 7];
    let mut field = FieldInfo::new(vec![1]);
    field.object_references = vec![7];
    info.field_infos.push(field);
    validate_storage_metadata(1, &info, &facts, &mut budget).expect("field subset");

    info.field_infos[0].object_references = vec![11];
    assert!(validate_storage_metadata(1, &info, &facts, &mut budget).is_err());
}
