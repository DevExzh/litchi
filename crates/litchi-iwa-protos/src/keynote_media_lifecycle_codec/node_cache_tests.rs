use super::{
    SlideNodeBuildCacheEdit, decode_slide_node_build_cache,
    decode_slide_node_build_cache_with_report, prepare_slide_node_build_cache_rewrite,
    rewrite_slide_node_build_cache_for_event_count, rewrite_slide_node_build_cache_with_report,
};
use crate::keynote_media_lifecycle_codec::{DecodeLimit, DecodeOptions};

fn varint(mut value: u64) -> Vec<u8> {
    let mut encoded = Vec::new();
    while value >= 0x80 {
        encoded.push((value as u8) | 0x80);
        value >>= 7;
    }
    encoded.push(value as u8);
    encoded
}

fn varint_field(field: u32, value: u64) -> Vec<u8> {
    let mut encoded = varint(u64::from(field) << 3);
    encoded.extend(varint(value));
    encoded
}

fn stale_node(
    count: u64,
    cache_version: u64,
    count_valid: u64,
    explicit: u64,
    explicit_version: u64,
    explicit_valid: u64,
) -> Vec<u8> {
    [
        varint_field(15, count),
        varint_field(26, cache_version),
        varint_field(22, count_valid),
        varint_field(20, explicit),
        varint_field(27, explicit_version),
        varint_field(23, explicit_valid),
        varint_field(100, 9),
    ]
    .concat()
}

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::for_source(source)
}

#[test]
fn scalar_snapshot_is_borrowed_and_lazy_projection_matches() {
    let source = stale_node(3, 2, 1, 1, 2, 1);
    let (snapshot, report) = decode_slide_node_build_cache_with_report(&source, options(&source))
        .expect("native cache fields should decode");
    assert_eq!(snapshot.source(), source.as_slice());
    assert_eq!(snapshot.build_event_count(), Some(3));
    assert_eq!(snapshot.build_event_count_cache_version(), Some(2));
    assert_eq!(snapshot.build_event_count_is_up_to_date(), Some(true));
    assert_eq!(snapshot.has_explicit_builds(), Some(true));
    assert_eq!(snapshot.has_explicit_builds_cache_version(), Some(2));
    assert_eq!(snapshot.has_explicit_builds_is_up_to_date(), Some(true));
    assert_eq!(report.input_bytes(), source.len());
    assert!(report.work_bytes() >= source.len().saturating_mul(2));
}

#[test]
fn already_invalidated_edit_is_byte_exact_and_reports_no_change() {
    let source = [
        varint_field(26, u64::from(u32::MAX)),
        varint_field(27, u64::from(u32::MAX)),
        varint_field(100, 9),
    ]
    .concat();
    let (output, report) = rewrite_slide_node_build_cache_with_report(
        &source,
        SlideNodeBuildCacheEdit::invalidate(),
        options(&source),
    )
    .expect("identity cache rewrite should succeed");
    assert_eq!(output, source);
    assert!(!report.changed());
    assert_eq!(report.output_bytes(), source.len());
}

#[test]
fn invalidation_removes_stale_cache_scalars_and_sets_unknown_versions() {
    let source = stale_node(7, 17, 1, 1, 99, 1);
    let unknown = varint_field(100, 9);
    let (output, report) = rewrite_slide_node_build_cache_with_report(
        &source,
        SlideNodeBuildCacheEdit::invalidate(),
        options(&source),
    )
    .expect("cache invalidation should succeed");
    let snapshot = decode_slide_node_build_cache(&output, options(&output))
        .expect("rewritten cache should decode");
    assert_eq!(snapshot.build_event_count(), None);
    assert_eq!(snapshot.build_event_count_cache_version(), Some(u32::MAX));
    assert_eq!(snapshot.build_event_count_is_up_to_date(), None);
    assert_eq!(snapshot.has_explicit_builds(), None);
    assert_eq!(snapshot.has_explicit_builds_cache_version(), Some(u32::MAX));
    assert_eq!(snapshot.has_explicit_builds_is_up_to_date(), None);
    assert!(output.windows(unknown.len()).any(|span| span == unknown));
    assert!(report.changed());
}

#[test]
fn absent_scalars_are_appended_in_native_patch_order() {
    let source = varint_field(100, 9);
    let prepared = prepare_slide_node_build_cache_rewrite(
        &source,
        SlideNodeBuildCacheEdit::invalidate(),
        options(&source)
            .with_max_output_bytes(32)
            .with_max_work_bytes(1024),
    )
    .expect("missing cache fields should be appendable");
    assert_eq!(prepared.snapshot().source(), source.as_slice());
    let (output, _) = prepared.commit().expect("append rewrite should commit");
    let expected_suffix = [
        varint_field(26, u64::from(u32::MAX)),
        varint_field(27, u64::from(u32::MAX)),
    ]
    .concat();
    assert_eq!(&output[..source.len()], source.as_slice());
    assert_eq!(&output[source.len()..], expected_suffix.as_slice());
}

#[test]
fn duplicate_wrong_wire_noncanonical_and_overflow_are_rejected() {
    let duplicate = [varint_field(15, 1), varint_field(15, 1)].concat();
    assert_eq!(
        decode_slide_node_build_cache(&duplicate, options(&duplicate))
            .expect_err("duplicate singular scalar must fail")
            .duplicate_field(),
        Some("KN.SlideNodeArchive.buildEventCount")
    );

    let wrong_wire = vec![0x7a, 0x00];
    assert!(decode_slide_node_build_cache(&wrong_wire, options(&wrong_wire)).is_err());

    let noncanonical_bool = varint_field(20, 2);
    assert!(
        decode_slide_node_build_cache(&noncanonical_bool, options(&noncanonical_bool)).is_err()
    );

    let overflow = varint_field(15, u64::from(u32::MAX) + 1);
    assert!(decode_slide_node_build_cache(&overflow, options(&overflow)).is_err());
}

#[test]
fn output_and_work_limits_are_charged_before_commit() {
    let source = varint_field(100, 9);
    let output_limited = options(&source).with_max_output_bytes(source.len());
    let error = prepare_slide_node_build_cache_rewrite(
        &source,
        SlideNodeBuildCacheEdit::invalidate(),
        output_limited,
    )
    .expect_err("append must exceed the output ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::OutputBytes { .. })
    ));

    let work_limited = options(&source).with_max_work_bytes(source.len());
    let error = decode_slide_node_build_cache(&source, work_limited)
        .expect_err("lazy parity must charge source work");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::Work { .. })
    ));
}

#[test]
fn event_count_rewrite_matches_native_creation_cache_rules() {
    let source = [
        varint_field(15, 9),
        varint_field(26, 99),
        varint_field(22, 1),
        varint_field(20, 0),
        varint_field(27, u64::from(u32::MAX)),
        varint_field(23, 0),
        varint_field(100, 9),
    ]
    .concat();
    let options = options(&source)
        .with_max_output_bytes(256)
        .with_max_work_bytes(4096);
    let output = rewrite_slide_node_build_cache_for_event_count(&source, 1, options)
        .expect("fresh build cache should be writable");
    let snapshot = decode_slide_node_build_cache(&output, options)
        .expect("rewritten fresh build cache should decode");
    assert_eq!(snapshot.build_event_count(), Some(1));
    assert_eq!(snapshot.build_event_count_cache_version(), Some(2));
    assert_eq!(snapshot.has_explicit_builds(), Some(true));
    assert_eq!(snapshot.has_explicit_builds_cache_version(), Some(2));
    assert_eq!(snapshot.build_event_count_is_up_to_date(), Some(true));
    assert_eq!(snapshot.has_explicit_builds_is_up_to_date(), Some(false));
    assert!(output.windows(2).any(|span| span == [0xa0, 0x06]));

    let output = rewrite_slide_node_build_cache_for_event_count(&output, 0, options)
        .expect("empty build cache should be writable");
    let snapshot = decode_slide_node_build_cache(&output, options)
        .expect("rewritten empty build cache should decode");
    assert_eq!(snapshot.build_event_count(), None);
    assert_eq!(snapshot.build_event_count_cache_version(), Some(u32::MAX));
    assert_eq!(snapshot.has_explicit_builds(), Some(false));
    assert_eq!(snapshot.has_explicit_builds_cache_version(), Some(2));
    assert_eq!(snapshot.build_event_count_is_up_to_date(), Some(true));
    assert_eq!(snapshot.has_explicit_builds_is_up_to_date(), Some(false));
    assert!(output.windows(2).any(|span| span == [0xa0, 0x06]));

    let output = rewrite_slide_node_build_cache_for_event_count(&[], 1, options)
        .expect("absent deprecated flags should remain absent");
    let snapshot =
        decode_slide_node_build_cache(&output, options).expect("fresh cache should decode");
    assert_eq!(snapshot.build_event_count_is_up_to_date(), None);
    assert_eq!(snapshot.has_explicit_builds_is_up_to_date(), None);
}
