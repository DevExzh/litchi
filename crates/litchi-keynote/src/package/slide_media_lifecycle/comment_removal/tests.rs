use std::path::PathBuf;

use litchi_iwa_common::{
    WireLimits, wire::append_length_delimited_field, wire::append_varint_field,
};
use litchi_iwa_core::MessageInfo;

use super::*;

const NATIVE_AUTHOR: u64 = 2_653_721;
const COMMENT_STORAGE: u64 = 2_653_723;
const COMMENT_COMPONENT: &str = "Index/Slide-2652150.iwa";
const WRONG_TYPE_TARGET: u64 = 2_652_149;

fn native_package() -> Package {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/media-comments-duplicate-native.key"),
    )
    .expect("native comment fixture");
    Package::from_bytes(&bytes).expect("native comment package")
}

#[test]
fn retained_comment_closure_handles_decreasing_storage_ids() {
    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let mut retained = vec![30];
    let mut edges = vec![(30, 20), (20, 10), (10, 5)];

    close_retained_comment_storage_ids(&mut retained, &mut edges, &mut budget)
        .expect("retained closure");

    assert_eq!(retained, [5, 10, 20, 30]);
}

#[test]
fn external_author_reference_does_not_mark_source_author_usage() {
    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let mut retained = Vec::new();
    let mut comment_edges = Vec::new();
    let mut selected_comment_author_edges = Vec::new();
    let mut used_author_ids = Vec::new();
    let comment_storage_ids = [];
    let author_ids = [NATIVE_AUTHOR];

    visit_reference(
        1,
        NATIVE_AUTHOR,
        false,
        false,
        true,
        &comment_storage_ids,
        &author_ids,
        &mut retained,
        &mut comment_edges,
        &mut selected_comment_author_edges,
        &mut used_author_ids,
        &mut budget,
    )
    .expect("source-component author reference");
    visit_reference(
        2,
        NATIVE_AUTHOR,
        false,
        false,
        false,
        &comment_storage_ids,
        &author_ids,
        &mut retained,
        &mut comment_edges,
        &mut selected_comment_author_edges,
        &mut used_author_ids,
        &mut budget,
    )
    .expect("external author reference");

    assert_eq!(used_author_ids, [NATIVE_AUTHOR]);
}

#[test]
fn movie_payload_reference_missing_from_header_is_rejected_atomically() {
    let mut drawable = Vec::new();
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, COMMENT_STORAGE).expect("comment identifier");
    append_length_delimited_field(&mut drawable, 6, &reference).expect("comment edge");
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable).expect("movie drawable");

    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let mut info = MessageInfo::new(
        MOVIE_MESSAGE_TYPE,
        payload.len().try_into().expect("movie payload length"),
    );

    assert!(
        validate_movie_payload_relationship(
            &package,
            COMMENT_COMPONENT,
            &payload,
            &info,
            &[COMMENT_STORAGE],
            WireLimits::default(),
            &mut budget,
        )
        .is_err()
    );
    info.object_references.push(COMMENT_STORAGE);
    assert!(
        validate_movie_payload_relationship(
            &package,
            COMMENT_COMPONENT,
            &payload,
            &info,
            &[COMMENT_STORAGE],
            WireLimits::default(),
            &mut budget,
        )
        .is_ok()
    );
}

#[test]
fn movie_comment_target_must_be_a_same_component_comment_storage() {
    let mut drawable = Vec::new();
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, WRONG_TYPE_TARGET).expect("target identifier");
    append_length_delimited_field(&mut drawable, 6, &reference).expect("comment edge");
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable).expect("movie drawable");

    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let mut info = MessageInfo::new(
        MOVIE_MESSAGE_TYPE,
        payload.len().try_into().expect("movie payload length"),
    );
    info.object_references.push(WRONG_TYPE_TARGET);

    assert!(
        validate_movie_payload_relationship(
            &package,
            COMMENT_COMPONENT,
            &payload,
            &info,
            &[WRONG_TYPE_TARGET],
            WireLimits::default(),
            &mut budget,
        )
        .is_err()
    );
}

#[test]
fn comment_reply_reference_missing_from_header_is_rejected() {
    let mut reply = Vec::new();
    append_varint_field(&mut reply, 1, 9_001).expect("reply identifier");
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 4, &reply).expect("reply edge");

    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let mut info = MessageInfo::new(
        COMMENT_STORAGE_MESSAGE_TYPE,
        payload.len().try_into().expect("comment payload length"),
    );

    assert!(
        validate_comment_storage_relationship(
            &package,
            COMMENT_COMPONENT,
            1,
            &info,
            &payload,
            WireLimits::default(),
            &mut budget,
        )
        .is_err()
    );
    info.object_references.push(9_001);
    assert!(
        validate_comment_storage_relationship(
            &package,
            COMMENT_COMPONENT,
            1,
            &info,
            &payload,
            WireLimits::default(),
            &mut budget,
        )
        .is_err()
    );
}

#[test]
fn opaque_comment_root_extension_is_preserved_by_relationship_census() {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 10_000, b"opaque-extension")
        .expect("opaque root extension");

    let package = native_package();
    let mut budget = LifecycleBudget::for_package(&package).expect("lifecycle budget");
    let info = MessageInfo::new(
        COMMENT_STORAGE_MESSAGE_TYPE,
        payload.len().try_into().expect("comment payload length"),
    );

    validate_comment_storage_relationship(
        &package,
        COMMENT_COMPONENT,
        1,
        &info,
        &payload,
        WireLimits::default(),
        &mut budget,
    )
    .expect("opaque root extension remains neutral");
}
