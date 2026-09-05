use super::*;

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(number) << 3);
    push_varint(&mut output, value);
    output
}

fn fixed32_field(number: u32, value: f32) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(number) << 3 | 5);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
    output
}

fn length_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(number) << 3 | 2);
    push_varint(&mut output, payload.len() as u64);
    output.extend_from_slice(payload);
    output
}

fn group_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(number) << 3 | 3);
    output.extend_from_slice(payload);
    push_varint(&mut output, u64::from(number) << 3 | 4);
    output
}

fn archive(payload: &[u8]) -> Vec<u8> {
    length_field(IMAGE_ADJUSTMENTS_FIELD, payload)
}

fn permissive(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(source.len(), usize::MAX, usize::MAX, MAX_RECURSION)
        .with_max_output_bytes(usize::MAX)
}

fn contains_bytes(source: &[u8], needle: &[u8]) -> bool {
    source.windows(needle.len()).any(|window| window == needle)
}

fn assert_limit(error: DecodeError, expected: DecodeLimit) {
    assert_eq!(error.limit_kind(), Some(expected), "{error}");
    let (observed, maximum) = error
        .limit_values()
        .expect("resource errors expose observed and maximum values");
    assert!(observed > maximum, "{observed} must exceed {maximum}");
}

fn rich_source() -> Vec<u8> {
    let advanced = fixed32_field(3, 0.4);
    let unknown_inner = varint_field(99, 990);
    let mut nested = Vec::new();
    nested.extend_from_slice(&advanced);
    nested.extend_from_slice(&unknown_inner);
    nested.extend_from_slice(&fixed32_field(EXPOSURE_FIELD, 0.0));
    nested.extend_from_slice(&fixed32_field(SATURATION_FIELD, 0.0));
    nested.extend_from_slice(&varint_field(ENHANCE_FIELD, 0));

    let mut source = archive(&nested);
    source.extend_from_slice(&varint_field(100, 1_000));
    source.extend_from_slice(&varint_field(100, 1_001));
    source
}

fn grouped_source() -> Vec<u8> {
    let inner_group = group_field(99, &varint_field(98, 7));
    let mut nested = inner_group;
    nested.extend_from_slice(&fixed32_field(EXPOSURE_FIELD, 0.0));

    let mut source = archive(&nested);
    source.extend_from_slice(&group_field(100, &varint_field(101, 1)));
    source
}

#[test]
fn parity_unknown_advanced_and_root_fields_survive_rewrites_and_clear() {
    let source = rich_source();
    let advanced = fixed32_field(3, 0.4);
    let unknown_inner = varint_field(99, 990);
    let unknown_root = varint_field(100, 1_000);
    let unknown_root_duplicate = varint_field(100, 1_001);

    let snapshot = decode_image_adjustments(&source, permissive(&source)).expect("rich source");
    assert_eq!(snapshot.exposure(), Some(0.0));
    assert_eq!(snapshot.saturation(), Some(0.0));
    assert_eq!(snapshot.enhance(), Some(false));
    assert!(snapshot.has_image_adjustments());

    let changed = rewrite_image_adjustments(
        &source,
        ImageAdjustmentsWrite::from_values(Some(0.25), Some(-0.5), Some(true)),
        permissive(&source),
    )
    .expect("selected rewrite");
    assert!(contains_bytes(&changed, &advanced));
    assert!(contains_bytes(&changed, &unknown_inner));
    assert!(contains_bytes(&changed, &unknown_root));
    assert!(contains_bytes(&changed, &unknown_root_duplicate));
    let changed_snapshot = decode_image_adjustments(&changed, permissive(&changed)).unwrap();
    assert_eq!(changed_snapshot.exposure(), Some(0.25));
    assert_eq!(changed_snapshot.saturation(), Some(-0.5));
    assert_eq!(changed_snapshot.enhance(), Some(true));

    let restored = rewrite_image_adjustments(
        &changed,
        ImageAdjustmentsWrite::from_values(Some(0.0), Some(0.0), Some(false)),
        permissive(&changed),
    )
    .expect("inverse rewrite");
    assert_eq!(restored, source);

    let cleared =
        rewrite_image_adjustments(&changed, ImageAdjustmentsWrite::new(), permissive(&changed))
            .expect("clear selected controls");
    assert!(contains_bytes(&cleared, &advanced));
    assert!(contains_bytes(&cleared, &unknown_inner));
    assert!(contains_bytes(&cleared, &unknown_root));
    assert!(contains_bytes(&cleared, &unknown_root_duplicate));
    let cleared_snapshot = decode_image_adjustments(&cleared, permissive(&cleared)).unwrap();
    assert_eq!(cleared_snapshot.exposure(), None);
    assert_eq!(cleared_snapshot.saturation(), None);
    assert_eq!(cleared_snapshot.enhance(), None);
    assert!(cleared_snapshot.has_image_adjustments());
}

#[test]
fn parity_unknown_groups_are_opaque_and_depth_is_accounted() {
    let source = grouped_source();
    let root_group = group_field(100, &varint_field(101, 1));
    let inner_group = group_field(99, &varint_field(98, 7));

    let (snapshot, report) = decode_image_adjustments_with_report(
        &source,
        DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 3),
    )
    .expect("balanced groups");
    assert_eq!(snapshot.exposure(), Some(0.0));
    assert!(report.max_depth() >= 3);

    let no_op = rewrite_image_adjustments(
        &source,
        ImageAdjustmentsWrite::from_values(Some(0.0), None, None),
        DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 3)
            .with_max_output_bytes(source.len()),
    )
    .expect("group-bearing no-op");
    assert_eq!(no_op, source);

    let changed = rewrite_image_adjustments(
        &source,
        ImageAdjustmentsWrite::from_values(Some(0.25), None, None),
        permissive(&source),
    )
    .expect("group-bearing rewrite");
    assert!(contains_bytes(&changed, &root_group));
    assert!(contains_bytes(&changed, &inner_group));
    assert_eq!(
        decode_image_adjustments(&changed, permissive(&changed))
            .unwrap()
            .exposure(),
        Some(0.25)
    );

    let shallow = DecodeOptions::new(source.len(), usize::MAX, usize::MAX, 2);
    assert_limit(
        decode_image_adjustments(&source, shallow).expect_err("inner group depth"),
        DecodeLimit::Nesting,
    );
    assert_limit(
        rewrite_image_adjustments(
            &source,
            ImageAdjustmentsWrite::from_values(Some(0.0), None, None),
            shallow,
        )
        .expect_err("rewrite depth"),
        DecodeLimit::Nesting,
    );
}

#[test]
fn parity_canonical_varints_and_booleans_are_strict() {
    let canonical = archive(&varint_field(ENHANCE_FIELD, 1));
    assert_eq!(
        decode_image_adjustments(&canonical, permissive(&canonical))
            .unwrap()
            .enhance(),
        Some(true)
    );

    let noncanonical_unknown_key = [0xa0, 0x86, 0x00, 0x01];
    assert!(
        decode_image_adjustments(
            &noncanonical_unknown_key,
            permissive(&noncanonical_unknown_key)
        )
        .is_err()
    );

    let noncanonical_unknown_value = [0xa0, 0x06, 0x80, 0x00];
    assert!(
        decode_image_adjustments(
            &noncanonical_unknown_value,
            permissive(&noncanonical_unknown_value)
        )
        .is_err()
    );

    for value in [2, 127, u64::MAX] {
        let nested = varint_field(ENHANCE_FIELD, value);
        let source = archive(&nested);
        assert!(
            decode_image_adjustments(&source, permissive(&source)).is_err(),
            "noncanonical bool {value} was accepted"
        );
    }
    let nested = [0x68, 0x80, 0x00];
    let source = archive(&nested);
    assert!(decode_image_adjustments(&source, permissive(&source)).is_err());
}

#[test]
fn parity_duplicate_outer_and_selected_fields_are_rejected() {
    let payload = fixed32_field(EXPOSURE_FIELD, 0.25);
    let mut duplicate_outer = archive(&payload);
    duplicate_outer.extend_from_slice(&archive(&payload));
    let error = decode_image_adjustments(&duplicate_outer, permissive(&duplicate_outer))
        .expect_err("duplicate outer field");
    assert_eq!(error.limit_kind(), None);

    let mut duplicate_inner = payload.clone();
    duplicate_inner.extend_from_slice(&fixed32_field(EXPOSURE_FIELD, 0.25));
    let duplicate_inner = archive(&duplicate_inner);
    assert!(decode_image_adjustments(&duplicate_inner, permissive(&duplicate_inner)).is_err());
}

#[test]
fn parity_optionals_distinguish_absent_empty_and_explicit_neutral_presence() {
    let omitted = decode_image_adjustments(&[], DecodeOptions::new(0, 0, 0, 2)).unwrap();
    assert_eq!(omitted.exposure(), None);
    assert_eq!(omitted.saturation(), None);
    assert_eq!(omitted.enhance(), None);
    assert!(!omitted.has_image_adjustments());

    let empty = [0x72, 0x00];
    let empty_snapshot = decode_image_adjustments(&empty, permissive(&empty)).unwrap();
    assert_eq!(empty_snapshot.exposure(), None);
    assert!(empty_snapshot.has_image_adjustments());
    assert_eq!(
        rewrite_image_adjustments(&empty, ImageAdjustmentsWrite::new(), permissive(&empty))
            .unwrap(),
        empty
    );

    let explicit = rewrite_image_adjustments(
        &[],
        ImageAdjustmentsWrite::from_values(Some(0.0), Some(0.0), Some(false)),
        DecodeOptions::new(0, usize::MAX, usize::MAX, 2).with_max_output_bytes(64),
    )
    .expect("explicit neutral controls");
    let explicit_snapshot = decode_image_adjustments(&explicit, permissive(&explicit)).unwrap();
    assert_eq!(explicit_snapshot.exposure(), Some(0.0));
    assert_eq!(explicit_snapshot.saturation(), Some(0.0));
    assert_eq!(explicit_snapshot.enhance(), Some(false));
    assert!(explicit_snapshot.has_image_adjustments());
    assert!(!explicit.is_empty());

    let removed = rewrite_image_adjustments(
        &explicit,
        ImageAdjustmentsWrite::new(),
        permissive(&explicit),
    )
    .expect("clear explicit controls");
    assert!(removed.is_empty());
    let removed_snapshot = decode_image_adjustments(&removed, DecodeOptions::new(0, 0, 0, 2))
        .expect("omitted after clear");
    assert!(!removed_snapshot.has_image_adjustments());
}

#[test]
fn parity_negative_zero_is_preserved_by_semantic_noops() {
    let mut nested = fixed32_field(EXPOSURE_FIELD, -0.0);
    nested.extend_from_slice(&fixed32_field(SATURATION_FIELD, 0.0));
    nested.extend_from_slice(&varint_field(ENHANCE_FIELD, 0));
    let source = archive(&nested);
    let snapshot = decode_image_adjustments(&source, permissive(&source)).unwrap();
    assert_eq!(snapshot.exposure().unwrap().to_bits(), (-0.0f32).to_bits());

    let positive_request = ImageAdjustmentsWrite::from_values(Some(0.0), Some(0.0), Some(false));
    assert_eq!(
        rewrite_image_adjustments(&source, positive_request, permissive(&source)).unwrap(),
        source
    );

    let positive_source = archive(&{
        let mut payload = fixed32_field(EXPOSURE_FIELD, 0.0);
        payload.extend_from_slice(&fixed32_field(SATURATION_FIELD, -0.0));
        payload.extend_from_slice(&varint_field(ENHANCE_FIELD, 0));
        payload
    });
    let negative_request = ImageAdjustmentsWrite::from_values(Some(-0.0), Some(-0.0), Some(false));
    assert_eq!(
        rewrite_image_adjustments(
            &positive_source,
            negative_request,
            permissive(&positive_source)
        )
        .unwrap(),
        positive_source
    );
}

#[test]
fn parity_invalid_selected_values_and_framing_fail_closed() {
    let malformed = [
        vec![0x72, 0x02, 0x08, 0x01], // selected exposure with varint wire type
        vec![0x72, 0x02, 0x08, 0x80], // unknown selected payload is truncated
        vec![0xf0, 0x86, 0x00, 0x00], // noncanonical unknown field key
    ];
    for source in malformed {
        assert!(
            decode_image_adjustments(&source, permissive(&source)).is_err(),
            "malformed source unexpectedly decoded: {source:?}"
        );
    }

    for value in [f32::NAN, f32::INFINITY, f32::NEG_INFINITY, -1.01, 1.01] {
        let result = prepare_image_adjustments_rewrite(
            &[],
            ImageAdjustmentsWrite::from_values(Some(value), None, None),
            DecodeOptions::new(0, usize::MAX, usize::MAX, 2).with_max_output_bytes(64),
        );
        assert!(result.is_err(), "invalid adjustment {value:?} accepted");
    }
}

#[test]
fn parity_decode_limits_are_inclusive_and_minus_one_is_rejected() {
    let source = rich_source();
    let (_, report) = decode_image_adjustments_with_report(&source, permissive(&source))
        .expect("baseline report");

    let exact = DecodeOptions::new(
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
    )
    .with_max_output_bytes(usize::MAX);
    assert!(decode_image_adjustments(&source, exact).is_ok());

    assert_limit(
        decode_image_adjustments(&source, exact.with_max_input_bytes(source.len() - 1))
            .expect_err("input bytes below exact source length"),
        DecodeLimit::InputBytes,
    );
    assert_limit(
        decode_image_adjustments(
            &source,
            exact.with_resource_limits(report.fields() - 1, report.work_bytes()),
        )
        .expect_err("fields below exact visit count"),
        DecodeLimit::Fields,
    );
    assert_limit(
        decode_image_adjustments(
            &source,
            exact.with_resource_limits(report.fields(), report.work_bytes() - 1),
        )
        .expect_err("work below exact accounting"),
        DecodeLimit::Work,
    );
    assert_limit(
        decode_image_adjustments(&source, exact.with_recursion_limit(report.max_depth() - 1))
            .expect_err("nesting below exact depth"),
        DecodeLimit::Nesting,
    );
}

#[test]
fn parity_rewrite_checks_aggregate_source_and_candidate_limits() {
    let source = rich_source();
    let write = ImageAdjustmentsWrite::from_values(Some(0.25), Some(-0.5), Some(true));
    let prepared = prepare_image_adjustments_rewrite(&source, write, permissive(&source))
        .expect("baseline preparation");
    let requirements = prepared.execution_requirements();
    assert!(requirements.fields > 0);
    assert!(requirements.work_bytes > 0);
    assert!(requirements.max_depth > 0);

    let exact = DecodeOptions::new(
        source.len(),
        requirements.fields,
        requirements.work_bytes,
        requirements.max_depth,
    )
    .with_max_output_bytes(requirements.output_bytes);
    let output = rewrite_image_adjustments(&source, write, exact).expect("exact caller limits");
    assert_eq!(output.len(), requirements.output_bytes);

    assert_limit(
        rewrite_image_adjustments(
            &source,
            write,
            exact.with_max_output_bytes(requirements.output_bytes - 1),
        )
        .expect_err("output below candidate size"),
        DecodeLimit::OutputBytes,
    );
    assert_limit(
        rewrite_image_adjustments(
            &source,
            write,
            exact.with_resource_limits(requirements.fields - 1, requirements.work_bytes),
        )
        .expect_err("fields below aggregate source plus candidate"),
        DecodeLimit::Fields,
    );
    assert_limit(
        rewrite_image_adjustments(
            &source,
            write,
            exact.with_resource_limits(requirements.fields, requirements.work_bytes - 1),
        )
        .expect_err("work below aggregate source plus candidate"),
        DecodeLimit::Work,
    );
    assert_limit(
        rewrite_image_adjustments(
            &source,
            write,
            exact.with_recursion_limit(requirements.max_depth - 1),
        )
        .expect_err("nesting below aggregate candidate depth"),
        DecodeLimit::Nesting,
    );

    let empty_fields = DecodeOptions::new(0, 0, usize::MAX, 2).with_max_output_bytes(64);
    assert_limit(
        rewrite_image_adjustments(
            &[],
            ImageAdjustmentsWrite::from_values(Some(0.25), None, None),
            empty_fields,
        )
        .expect_err("missing control exceeds zero field budget"),
        DecodeLimit::Fields,
    );
    let empty_work = DecodeOptions::new(0, usize::MAX, 0, 2).with_max_output_bytes(64);
    assert_limit(
        rewrite_image_adjustments(
            &[],
            ImageAdjustmentsWrite::from_values(Some(0.25), None, None),
            empty_work,
        )
        .expect_err("missing control exceeds zero work budget"),
        DecodeLimit::Work,
    );

    let tight_source = varint_field(100, 7);
    let (_, tight_report) =
        decode_image_adjustments_with_report(&tight_source, permissive(&tight_source))
            .expect("tight source report");
    let tight_options =
        DecodeOptions::new(tight_source.len(), tight_report.fields(), usize::MAX, 2)
            .with_max_output_bytes(64);
    assert_limit(
        rewrite_image_adjustments(
            &tight_source,
            ImageAdjustmentsWrite::from_values(Some(0.25), None, None),
            tight_options,
        )
        .expect_err("appending a missing control exceeds the tight field budget"),
        DecodeLimit::Fields,
    );
}

#[test]
fn parity_prepared_limits_are_inclusive_and_allocation_bounded() {
    let source = rich_source();
    let write = ImageAdjustmentsWrite::from_values(Some(0.25), Some(-0.5), Some(true));
    let prepared =
        prepare_image_adjustments_rewrite(&source, write, permissive(&source)).expect("prepare");
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(requirements.exact())
        .expect("exact prepared execution");
    assert_eq!(output.output().len(), requirements.output_bytes);
    assert_eq!(output.report().fields(), requirements.fields);
    assert_eq!(output.report().work_bytes(), requirements.work_bytes);
    assert_eq!(output.report().max_depth(), requirements.max_depth);

    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_output_bytes(requirements.output_bytes - 1),
            )
            .expect_err("prepared output ceiling"),
        DecodeLimit::OutputBytes,
    );
    assert_limit(
        prepared
            .execute(requirements.exact().with_fields(requirements.fields - 1))
            .expect_err("prepared field ceiling"),
        DecodeLimit::Fields,
    );
    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_work_bytes(requirements.work_bytes - 1),
            )
            .expect_err("prepared work ceiling"),
        DecodeLimit::Work,
    );
    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_max_depth(requirements.max_depth - 1),
            )
            .expect_err("prepared nesting ceiling"),
        DecodeLimit::Nesting,
    );
    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_allocations(requirements.allocations - 1),
            )
            .expect_err("prepared allocation ceiling"),
        DecodeLimit::Allocations,
    );
    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_retained_bytes(requirements.retained_bytes - 1),
            )
            .expect_err("prepared retained-byte ceiling"),
        DecodeLimit::Retained,
    );
    assert_limit(
        prepared
            .execute(
                requirements
                    .exact()
                    .with_scratch_bytes(requirements.scratch_bytes - 1),
            )
            .expect_err("prepared scratch ceiling"),
        DecodeLimit::Scratch,
    );
}

#[test]
fn parity_large_opaque_length_delimited_fields_charge_vector_scratch_only() {
    let large_source = length_field(100, &[0; 4096]);
    let small_source = varint_field(100, 0);

    let (_, large_report) = decode_image_adjustments_with_report(
        &large_source,
        DecodeOptions::new(large_source.len(), 1, large_source.len(), 8),
    )
    .expect("large opaque field");
    let (_, small_report) = decode_image_adjustments_with_report(
        &small_source,
        DecodeOptions::new(small_source.len(), 1, small_source.len(), 8),
    )
    .expect("small opaque field");

    assert_eq!(large_report.fields(), 1);
    assert_eq!(small_report.fields(), 1);
    assert_eq!(
        large_report.scratch_bytes(),
        small_report.scratch_bytes(),
        "opaque payload bytes belong to work, not ParsedField-vector scratch"
    );
    assert_eq!(large_report.scratch_bytes(), size_of::<ParsedField>());
    assert!(large_report.work_bytes() > small_report.work_bytes());
}

#[test]
fn parity_many_small_fields_keep_limits_and_noop_add_remove_accounting_exact() {
    let mut source = Vec::new();
    for field in 100..140 {
        source.extend_from_slice(&varint_field(field, u64::from(field - 100)));
    }

    let (_, report) = decode_image_adjustments_with_report(&source, permissive(&source))
        .expect("many opaque scalar fields");
    assert_eq!(report.fields(), 40);
    let exact_decode = DecodeOptions::new(
        source.len(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
    )
    .with_max_output_bytes(source.len());
    assert!(decode_image_adjustments(&source, exact_decode).is_ok());
    assert_limit(
        decode_image_adjustments(
            &source,
            exact_decode.with_resource_limits(report.fields() - 1, report.work_bytes()),
        )
        .expect_err("field budget below the forty visited fields"),
        DecodeLimit::Fields,
    );
    assert_limit(
        decode_image_adjustments(
            &source,
            exact_decode.with_resource_limits(report.fields(), report.work_bytes() - 1),
        )
        .expect_err("work budget below the exact scalar spans"),
        DecodeLimit::Work,
    );

    let no_op_write = ImageAdjustmentsWrite::new();
    let no_op_prepared =
        prepare_image_adjustments_rewrite(&source, no_op_write, permissive(&source))
            .expect("prepare many-field no-op");
    let no_op_requirements = no_op_prepared.execution_requirements();
    assert!(no_op_requirements.fields > report.fields());
    assert!(no_op_requirements.work_bytes > report.work_bytes());
    let no_op = rewrite_image_adjustments(
        &source,
        no_op_write,
        DecodeOptions::new(
            source.len(),
            no_op_requirements.fields,
            no_op_requirements.work_bytes,
            no_op_requirements.max_depth,
        )
        .with_max_output_bytes(no_op_requirements.output_bytes),
    )
    .expect("many-field no-op at aggregate exact limits");
    assert_eq!(no_op, source);

    let add_write = ImageAdjustmentsWrite::from_values(Some(0.25), None, None);
    let add_prepared = prepare_image_adjustments_rewrite(&source, add_write, permissive(&source))
        .expect("add missing selected control");
    let add_requirements = add_prepared.execution_requirements();
    let added = rewrite_image_adjustments(
        &source,
        add_write,
        DecodeOptions::new(
            source.len(),
            add_requirements.fields,
            add_requirements.work_bytes,
            add_requirements.max_depth,
        )
        .with_max_output_bytes(add_requirements.output_bytes),
    )
    .expect("add missing selected control at exact limits");
    let added_snapshot = decode_image_adjustments(&added, permissive(&added)).unwrap();
    assert_eq!(added_snapshot.exposure(), Some(0.25));
    for field in 100..140 {
        assert!(contains_bytes(
            &added,
            &varint_field(field, u64::from(field - 100))
        ));
    }

    let remove_write = ImageAdjustmentsWrite::new();
    let remove_prepared =
        prepare_image_adjustments_rewrite(&added, remove_write, permissive(&added))
            .expect("remove selected control");
    let remove_requirements = remove_prepared.execution_requirements();
    let removed = rewrite_image_adjustments(
        &added,
        remove_write,
        DecodeOptions::new(
            added.len(),
            remove_requirements.fields,
            remove_requirements.work_bytes,
            remove_requirements.max_depth,
        )
        .with_max_output_bytes(remove_requirements.output_bytes),
    )
    .expect("remove selected control at exact limits");
    assert_eq!(removed, source);
}
