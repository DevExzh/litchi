use super::*;

fn varint(mut value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
    output
}

fn varint_field(field: u32, value: u64) -> Vec<u8> {
    let mut output = varint(u64::from(field) << 3);
    output.extend(varint(value));
    output
}

fn bytes_field(field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = varint((u64::from(field) << 3) | 2);
    output.extend(varint(payload.len() as u64));
    output.extend_from_slice(payload);
    output
}

fn reference(identifier: u64) -> Vec<u8> {
    varint_field(1, identifier)
}

fn reference_field(field: u32, identifier: u64) -> Vec<u8> {
    bytes_field(field, &reference(identifier))
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    [varint_field(1, lower), varint_field(2, upper)].concat()
}

fn slide() -> Vec<u8> {
    [
        reference_field(1, 100),
        reference_field(2, 200),
        bytes_field(4, &bytes_field(2, &[])),
        reference_field(7, 300),
        reference_field(7, 301),
        varint_field(19, 1),
        reference_field(42, 301),
        reference_field(42, 300),
        reference_field(43, 400),
        bytes_field(100, b"preserve me"),
    ]
    .concat()
}

fn build() -> Vec<u8> {
    [
        reference_field(1, 300),
        bytes_field(2, b"byObject"),
        bytes_field(4, &bytes_field(77, b"opaque")),
        bytes_field(100, b"unknown"),
    ]
    .concat()
}

fn chunk() -> Vec<u8> {
    let source_uuid = uuid(11, 22);
    let identifier = [bytes_field(1, &source_uuid), varint_field(2, 4)].concat();
    [
        reference_field(1, 200),
        bytes_field(7, &identifier),
        bytes_field(8, &source_uuid),
        bytes_field(99, b"opaque"),
    ]
    .concat()
}

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::for_source(source)
        .with_max_message_bytes(source.len())
        .with_max_output_bytes(source.len().saturating_mul(4))
}

#[test]
fn slide_streams_all_lifecycle_lists_without_generated_repeated_views() {
    let source = slide();
    let snapshot = decode_slide_lifecycle(&source, options(&source)).expect("valid slide");
    assert_eq!(snapshot.style().identifier(), 100);
    assert!(snapshot.in_document());
    assert_eq!(
        snapshot
            .owned_drawables()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [300, 301]
    );
    assert_eq!(
        snapshot
            .drawables_z_order()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [301, 300]
    );
    assert_eq!(
        snapshot
            .builds()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [200]
    );
    assert_eq!(
        snapshot
            .build_chunks()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [400]
    );
}

#[test]
fn slide_rewrite_preserves_unknown_spans_and_has_source_witness() {
    let source = slide();
    let remaps = [IdentifierRewrite::new(300, 500)];
    let appended = [600];
    let edit = SlideLifecycleEdit::empty()
        .with_identifier_remaps(&remaps)
        .with_owned_drawables(&appended)
        .with_drawables_z_order(&appended);
    let prepared =
        prepare_slide_lifecycle_rewrite(&source, edit, options(&source)).expect("preparation");
    assert_eq!(prepared.source(), source.as_slice());
    let (output, report) = prepared.commit().expect("commit");
    assert!(report.changed());
    assert!(
        output
            .windows(b"preserve me".len())
            .any(|window| window == b"preserve me")
    );
    let snapshot = decode_slide_lifecycle(&output, options(&output)).expect("readback");
    assert_eq!(
        snapshot
            .owned_drawables()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [500, 301, 600]
    );
    assert_eq!(
        snapshot
            .drawables_z_order()
            .map(|reference| reference.identifier())
            .collect::<Vec<_>>(),
        [301, 500, 600]
    );
}

#[test]
fn build_and_chunk_rewrites_preserve_opaque_fields() {
    let build_source = build();
    let (build_output, build_report) = rewrite_build_with_report(
        &build_source,
        BuildLifecycleEdit::drawable(IdentifierRewrite::new(300, 500)),
        options(&build_source),
    )
    .expect("build rewrite");
    assert!(build_report.changed());
    assert!(
        build_output
            .windows(b"unknown".len())
            .any(|window| window == b"unknown")
    );
    assert_eq!(
        decode_build(&build_output, options(&build_output))
            .expect("build readback")
            .drawable()
            .expect("drawable")
            .identifier(),
        500
    );

    let chunk_source = chunk();
    let (chunk_output, chunk_report) = rewrite_build_chunk_with_report(
        &chunk_source,
        BuildChunkLifecycleEdit::empty()
            .with_build(IdentifierRewrite::new(200, 700))
            .with_uuid(UuidRewrite::new(Uuid::new(11, 22), Uuid::new(33, 44))),
        options(&chunk_source),
    )
    .expect("chunk rewrite");
    assert!(chunk_report.changed());
    assert!(
        chunk_output
            .windows(b"opaque".len())
            .any(|window| window == b"opaque")
    );
    let snapshot = decode_build_chunk(&chunk_output, options(&chunk_output)).expect("readback");
    assert_eq!(snapshot.build().identifier(), 700);
    assert_eq!(
        snapshot.chunk_identifier().expect("identifier").uuid(),
        Uuid::new(33, 44)
    );
    assert_eq!(
        snapshot.build_id().expect("build id").uuid(),
        Uuid::new(33, 44)
    );
}

#[test]
fn identity_build_and_chunk_rewrites_report_unchanged_bytes() {
    let build_source = build();
    let (build_output, build_report) = rewrite_build_with_report(
        &build_source,
        BuildLifecycleEdit::drawable(IdentifierRewrite::new(300, 300)),
        options(&build_source),
    )
    .expect("identity build rewrite");
    assert_eq!(build_output, build_source);
    assert!(!build_report.changed());

    let chunk_source = chunk();
    let (chunk_output, chunk_report) = rewrite_build_chunk_with_report(
        &chunk_source,
        BuildChunkLifecycleEdit::empty()
            .with_build(IdentifierRewrite::new(200, 200))
            .with_uuid(UuidRewrite::new(Uuid::new(11, 22), Uuid::new(11, 22))),
        options(&chunk_source),
    )
    .expect("identity chunk rewrite");
    assert_eq!(chunk_output, chunk_source);
    assert!(!chunk_report.changed());
}

#[test]
fn duplicate_and_deprecated_slide_topologies_fail_closed() {
    let source = slide();
    let mut duplicate = source.clone();
    duplicate.extend(reference_field(7, 300));
    assert!(decode_slide_lifecycle(&duplicate, options(&duplicate)).is_err());

    let mut deprecated = source;
    deprecated.extend(bytes_field(3, &[]));
    let error = decode_slide_lifecycle(&deprecated, options(&deprecated))
        .expect_err("deprecated inline chunks");
    assert_eq!(
        error.unsupported_reason(),
        Some("deprecated inline slide build chunks")
    );
}

#[test]
fn finite_limits_are_checked_before_rewrite_output() {
    let source = slide();
    let edit = SlideLifecycleEdit::empty().with_owned_drawables(&[700]);
    let error = prepare_slide_lifecycle_rewrite(
        &source,
        edit,
        options(&source).with_max_output_bytes(source.len()),
    )
    .expect_err("output ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::OutputBytes { .. })
    ));
}

#[test]
fn stale_removal_and_remap_target_append_are_rejected() {
    let source = slide();

    let stale = [999];
    let error = prepare_slide_lifecycle_rewrite(
        &source,
        SlideLifecycleEdit::empty().with_removed_identifiers(&stale),
        options(&source),
    )
    .expect_err("stale removal");
    assert_eq!(
        error.unsupported_reason(),
        Some("removed identifier is absent from every slide lifecycle list")
    );

    let remaps = [IdentifierRewrite::new(300, 500)];
    let appended = [500];
    let error = prepare_slide_lifecycle_rewrite(
        &source,
        SlideLifecycleEdit::empty()
            .with_identifier_remaps(&remaps)
            .with_owned_drawables(&appended),
        options(&source),
    )
    .expect_err("remap target append collision");
    assert_eq!(
        error.unsupported_reason(),
        Some("appended identifier conflicts with a remap target")
    );
}

#[test]
fn remap_edit_length_is_capped_before_quadratic_validation() {
    let source = slide();
    let remaps = [
        IdentifierRewrite::new(300, 500),
        IdentifierRewrite::new(301, 501),
        IdentifierRewrite::new(200, 600),
    ];
    let error = prepare_slide_lifecycle_rewrite(
        &source,
        SlideLifecycleEdit::empty().with_identifier_remaps(&remaps),
        options(&source).with_max_references(2),
    )
    .expect_err("remap list ceiling");
    assert!(matches!(
        error.resource_limit(),
        Some(DecodeLimit::References {
            observed: 3,
            maximum: 2
        })
    ));
}

#[test]
fn slide_lookup_work_is_precharged_before_source_loops() {
    let source = slide();
    let snapshot = decode_slide_lifecycle(&source, options(&source)).expect("valid slide");
    let remaps = [
        IdentifierRewrite::new(300, 500),
        IdentifierRewrite::new(301, 501),
    ];
    let edit = SlideLifecycleEdit::empty().with_identifier_remaps(&remaps);
    let constrained = DecodeOptions::new(
        source.len(),
        source.len().saturating_mul(4),
        source.len().saturating_mul(8),
        10,
        source.len().saturating_mul(2),
        16,
    );
    let mut budget = Budget::new(&source, constrained);
    budget.work_bytes = 0;
    let error = validate_slide_edit(&snapshot, edit, constrained, &mut budget)
        .expect_err("bounded lookup work");
    assert_eq!(
        error.resource_limit(),
        Some(DecodeLimit::Work {
            observed: 12,
            maximum: 10,
        })
    );
}

#[test]
fn rewrite_report_adds_emission_and_readback_work() {
    let source = [0_u8];
    let mut budget = Budget::new(&source, DecodeOptions::for_source(&source));
    budget.work_bytes = 10;
    let readback = DecodeReport {
        input_bytes: 1,
        fields: 0,
        work_bytes: 7,
        max_depth: 1,
        allocations: 0,
        retained_bytes: 1,
        scratch_bytes: 0,
    };
    let report = rewrite_report(&source, &budget, &readback, 1, true);
    assert_eq!(report.work_bytes(), 17);
}
