#![allow(
    clippy::expect_used,
    reason = "these integration tests use expect for concise public-API assertions"
)]

use litchi_xlsb::external_link::{
    DdeItem, ExternalLinkLimits, ExternalLinkLimitsError, ExternalLinkResource, Link,
    apply_with_limits, parse_external_link_model_with_limits, parse_external_link_with_limits,
    parse_external_link_with_relationship_with_limits, read_with_limits,
    write_external_link_stream_with_limits,
};

#[test]
fn default_limits_expose_documented_ceilings_and_aliases() {
    let limits = ExternalLinkLimits::DEFAULT;

    assert_eq!(ExternalLinkLimits::default(), limits);
    assert_eq!(
        ExternalLinkLimits::builder()
            .build()
            .expect("default policy"),
        limits
    );

    assert_eq!(limits.max_part_bytes(), ExternalLinkLimits::MAX_PART_BYTES);
    assert_eq!(
        limits.max_total_part_bytes(),
        ExternalLinkLimits::MAX_PART_BYTES
    );
    assert_eq!(
        limits.max_opaque_bytes(),
        ExternalLinkLimits::MAX_OPAQUE_BYTES
    );
    assert_eq!(
        limits.max_total_opaque_bytes(),
        ExternalLinkLimits::MAX_OPAQUE_BYTES
    );
    assert_eq!(
        limits.max_utf16_units(),
        ExternalLinkLimits::MAX_UTF16_UNITS
    );
    assert_eq!(
        limits.max_total_utf16_units(),
        ExternalLinkLimits::MAX_PART_BYTES / 2
    );
    assert_eq!(limits.max_records(), 1_048_576);
    assert_eq!(limits.max_cache_records(), 1_048_576);
    assert_eq!(limits.max_opaque_records(), 65_535);
    assert_eq!(limits.max_links(), 65_535);
    assert_eq!(limits.max_items(), 65_535);
    assert_eq!(limits.max_matrices(), 65_535);
    assert_eq!(limits.max_cells(), 1_048_576);
    assert_eq!(
        limits.max_decoded_semantic_bytes(),
        3 * (ExternalLinkLimits::MAX_PART_BYTES / 2)
    );
    assert_eq!(limits.max_retained_objects(), 1_441_786);

    assert_eq!(limits.max_total_records(), limits.max_records());
    assert_eq!(limits.max_total_cache_records(), limits.max_cache_records());
    assert_eq!(
        limits.max_total_opaque_records(),
        limits.max_opaque_records()
    );
    assert_eq!(limits.max_total_links(), limits.max_links());
    assert_eq!(limits.max_total_items(), limits.max_items());
    assert_eq!(limits.max_total_matrices(), limits.max_matrices());
    assert_eq!(limits.max_total_cells(), limits.max_cells());
    assert_eq!(
        limits.max_total_decoded_semantic_bytes(),
        limits.max_decoded_semantic_bytes()
    );
    assert_eq!(
        limits.max_total_retained_objects(),
        limits.max_retained_objects()
    );
}

#[test]
fn builder_supports_fluent_custom_limits_and_alias_setters() {
    let limits = ExternalLinkLimits::builder()
        .max_part_bytes(11)
        .max_total_part_bytes(12)
        .max_opaque_bytes(13)
        .max_total_opaque_bytes(14)
        .max_utf16_units(15)
        .max_total_utf16_units(16)
        .max_total_records(17)
        .max_total_cache_records(18)
        .max_total_opaque_records(19)
        .max_total_links(20)
        .max_total_items(21)
        .max_total_matrices(22)
        .max_total_cells(23)
        .max_total_decoded_semantic_bytes(24)
        .max_total_retained_objects(25)
        .build()
        .expect("custom policy");

    assert_eq!(limits.max_part_bytes(), 11);
    assert_eq!(limits.max_total_part_bytes(), 12);
    assert_eq!(limits.max_opaque_bytes(), 13);
    assert_eq!(limits.max_total_opaque_bytes(), 14);
    assert_eq!(limits.max_utf16_units(), 15);
    assert_eq!(limits.max_total_utf16_units(), 16);
    assert_eq!(limits.max_records(), 17);
    assert_eq!(limits.max_cache_records(), 18);
    assert_eq!(limits.max_opaque_records(), 19);
    assert_eq!(limits.max_links(), 20);
    assert_eq!(limits.max_items(), 21);
    assert_eq!(limits.max_matrices(), 22);
    assert_eq!(limits.max_cells(), 23);
    assert_eq!(limits.max_decoded_semantic_bytes(), 24);
    assert_eq!(limits.max_retained_objects(), 25);
}

#[test]
fn zero_per_part_and_aggregate_limits_are_valid() {
    let limits = ExternalLinkLimits::builder()
        .max_part_bytes(0)
        .max_total_part_bytes(0)
        .max_opaque_bytes(0)
        .max_total_opaque_bytes(0)
        .max_utf16_units(0)
        .max_total_utf16_units(0)
        .max_records(0)
        .max_cache_records(0)
        .max_opaque_records(0)
        .max_links(0)
        .max_items(0)
        .max_matrices(0)
        .max_cells(0)
        .max_decoded_semantic_bytes(0)
        .max_retained_objects(0)
        .build()
        .expect("zero policy");

    assert_eq!(limits.max_part_bytes(), 0);
    assert_eq!(limits.max_total_part_bytes(), 0);
    assert_eq!(limits.max_opaque_bytes(), 0);
    assert_eq!(limits.max_total_opaque_bytes(), 0);
    assert_eq!(limits.max_utf16_units(), 0);
    assert_eq!(limits.max_total_utf16_units(), 0);
    assert_eq!(limits.max_records(), 0);
    assert_eq!(limits.max_cache_records(), 0);
    assert_eq!(limits.max_opaque_records(), 0);
    assert_eq!(limits.max_links(), 0);
    assert_eq!(limits.max_items(), 0);
    assert_eq!(limits.max_matrices(), 0);
    assert_eq!(limits.max_cells(), 0);
    assert_eq!(limits.max_decoded_semantic_bytes(), 0);
    assert_eq!(limits.max_retained_objects(), 0);
}

#[test]
fn limit_builder_returns_typed_errors_with_exact_accessors() {
    let requested = ExternalLinkLimits::MAX_PART_BYTES + 1;
    let hard_maximum = ExternalLinkLimits::builder()
        .max_part_bytes(requested)
        .build()
        .expect_err("part hard maximum should be enforced");
    assert!(matches!(
        hard_maximum,
        ExternalLinkLimitsError::HardMaximum { .. }
    ));
    assert_eq!(hard_maximum.resource(), ExternalLinkResource::PartBytes);
    assert_eq!(hard_maximum.value(), requested);
    assert_eq!(hard_maximum.maximum(), ExternalLinkLimits::MAX_PART_BYTES);

    let per_part_exceeds_aggregate = ExternalLinkLimits::builder()
        .max_part_bytes(7)
        .max_total_part_bytes(6)
        .build()
        .expect_err("part aggregate should be enforced");
    assert!(matches!(
        per_part_exceeds_aggregate,
        ExternalLinkLimitsError::PerPartExceedsAggregate { .. }
    ));
    assert_eq!(
        per_part_exceeds_aggregate.resource(),
        ExternalLinkResource::PartBytes
    );
    assert_eq!(per_part_exceeds_aggregate.value(), 7);
    assert_eq!(per_part_exceeds_aggregate.maximum(), 6);
}

#[test]
fn public_with_limits_apis_round_trip_an_authored_dde_link() {
    let item = DdeItem::new("ITEM").expect("valid DDE item");
    let link = Link::dde_with_items("server", "topic", vec![item]).expect("valid DDE link");
    let limits = ExternalLinkLimits::DEFAULT;

    let encoded = write_external_link_stream_with_limits(&link, None, limits)
        .expect("write under default limits");
    let parsed = parse_external_link_with_limits(&encoded, limits).expect("parse under limits");
    assert_eq!(parsed.link(), &link);
    assert_eq!(
        parse_external_link_model_with_limits(&encoded, limits).expect("model parse under limits"),
        link
    );
    assert_eq!(
        parse_external_link_with_relationship_with_limits(&encoded, limits)
            .expect("relationship parse under limits")
            .link(),
        &link
    );

    let snapshot = read_with_limits(&encoded, limits).expect("read under limits");
    assert_eq!(snapshot.link(), &link);
    let patch = snapshot
        .edit()
        .commit()
        .expect("empty transaction")
        .patch()
        .clone();
    assert_eq!(
        apply_with_limits(&encoded, &patch, limits).expect("apply under limits"),
        encoded
    );
}
