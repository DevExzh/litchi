use super::{DecodeOptions, decode_modern, prepare_chart_data_rewrite};

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn varint_field(field: u32, value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3);
    push_varint(&mut output, value);
    output
}

fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 2);
    push_varint(
        &mut output,
        u64::try_from(payload.len()).expect("test payload length fits in a varint"),
    );
    output.extend_from_slice(payload);
    output
}

fn fixed64_field(field: u32, value: f64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 1);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
    output
}

fn group_tag(field: u32, wire_type: u64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | wire_type);
    output
}

fn unknown_group(field: u32, nested_fields: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    let mut output = group_tag(field, 3);
    output.extend(joined(nested_fields));
    output.extend(group_tag(field, 4));
    output
}

fn joined(fields: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    fields.into_iter().flatten().collect()
}

fn numeric(value: f64) -> Vec<u8> {
    fixed64_field(1, value)
}

fn date(value: f64) -> Vec<u8> {
    fixed64_field(4, value)
}

fn duration(value: f64) -> Vec<u8> {
    fixed64_field(3, value)
}

fn grid_row(values: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    length_field(
        3,
        &joined(values.into_iter().map(|value| length_field(1, &value))),
    )
}

fn grid(
    row_names: &[&str],
    column_names: &[&str],
    rows: impl IntoIterator<Item = Vec<Vec<u8>>>,
    id_map: &[u8],
) -> Vec<u8> {
    let mut output = Vec::new();
    for row_name in row_names {
        output.extend(length_field(1, row_name.as_bytes()));
    }
    for column_name in column_names {
        output.extend(length_field(2, column_name.as_bytes()));
    }
    for values in rows {
        output.extend(grid_row(values));
    }
    output.extend(length_field(4, id_map));
    output.extend(length_field(80, b"future grid extension"));
    output
}

fn modern(grid: &[u8]) -> Vec<u8> {
    let chart = joined([
        length_field(7, grid),
        varint_field(91, 0xfeed),
        length_field(92, b"future chart extension"),
    ]);
    joined([
        varint_field(4000, 7),
        length_field(10_000, &chart),
        length_field(93, b"future drawable extension"),
    ])
}

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::for_source(source).with_max_work_bytes(usize::MAX)
}

fn rich_source() -> Vec<u8> {
    modern(&grid(
        &["North", "South"],
        &["Q1", "Q2"],
        [
            vec![
                joined([
                    varint_field(90, 17),
                    numeric(17.25),
                    date(45_123.25),
                    duration(0.5),
                    length_field(89, b"cell future bytes"),
                ]),
                joined([length_field(88, b"blank future bytes"), date(45_124.0)]),
            ],
            vec![
                joined([numeric(-0.0), varint_field(87, 29)]),
                joined([duration(3.5), numeric(-9.25)]),
            ],
        ],
        &[0xff, 0x80, 0x80, 0x01],
    ))
}

fn values(snapshot: super::ChartDataSnapshot<'_>) -> Vec<Option<f64>> {
    snapshot
        .rows()
        .iter()
        .flat_map(|row| row.values())
        .collect()
}

fn grouped_source() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let value_group = unknown_group(13, [varint_field(4003, 0x77)]);
    let value = joined([value_group.clone(), date(45_125.5), numeric(17.25)]);
    let mut grid_source = grid(&["North"], &["Q1"], [vec![value]], &[0x31, 0x41]);
    let grid_group = unknown_group(
        12,
        [
            varint_field(4002, 0x66),
            unknown_group(14, [fixed64_field(4004, 12.5)]),
        ],
    );
    grid_source.extend_from_slice(&grid_group);
    let nested_root_group = unknown_group(11, [varint_field(4001, 0x55)]);
    let root_group = unknown_group(10, [nested_root_group.clone(), length_field(4005, b"root")]);
    let mut source = root_group.clone();
    source.extend_from_slice(&modern(&grid_source));
    (
        source,
        root_group,
        grid_group,
        value_group,
        nested_root_group,
    )
}

#[test]
fn rewrite_changes_only_numeric_scalars_and_preserves_all_other_wire_bytes() {
    let source = rich_source();
    let replacement = [vec![Some(1.5), None], vec![Some(-0.0), Some(7.75)]];
    let prepared = prepare_chart_data_rewrite(&source, &replacement, options(&source))
        .expect("same-shape numeric rewrite prepares");
    let output = prepared
        .execute(prepared.execution_requirements().exact())
        .expect("same-shape numeric rewrite executes");

    assert_ne!(output.output(), source.as_slice());
    assert!(output.report().changed());
    let output_bytes = output.output();
    for preserved in [
        b"North".as_slice(),
        b"South".as_slice(),
        b"Q1".as_slice(),
        b"Q2".as_slice(),
        b"cell future bytes".as_slice(),
        b"blank future bytes".as_slice(),
        b"future grid extension".as_slice(),
        b"future chart extension".as_slice(),
        b"future drawable extension".as_slice(),
        &[0xff, 0x80, 0x80, 0x01],
    ] {
        assert!(
            output_bytes
                .windows(preserved.len())
                .any(|window| window == preserved),
            "candidate dropped preserved bytes: {preserved:?}"
        );
    }
    assert!(
        output_bytes
            .windows(8)
            .any(|window| window == 1.5f64.to_bits().to_le_bytes())
    );
    assert!(
        output_bytes
            .windows(8)
            .any(|window| window == 7.75f64.to_bits().to_le_bytes())
    );

    let snapshot = decode_modern(output.output(), &options(output.output())).expect("readback");
    assert_eq!(values(snapshot), [Some(1.5), None, Some(-0.0), Some(7.75)]);
    assert_eq!(
        snapshot.row_labels().iter().collect::<Vec<_>>(),
        ["North", "South"]
    );
    assert_eq!(
        snapshot.column_labels().iter().collect::<Vec<_>>(),
        ["Q1", "Q2"]
    );
}

#[test]
fn signed_zero_noop_is_byte_exact_and_returns_owned_buffer() {
    let source = rich_source();
    let values = [vec![Some(17.25), None], vec![Some(-0.0), Some(-9.25)]];
    let prepared = prepare_chart_data_rewrite(&source, &values, options(&source))
        .expect("identical values prepare");
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(requirements.exact())
        .expect("identical values execute");

    assert_eq!(output.output(), source.as_slice());
    assert!(!output.report().changed());
    assert_eq!(requirements.output_bytes, source.len());
    // The candidate is byte-exact, but `execute` still returns an owned
    // buffer. Package transactions avoid this allocation by short-circuiting
    // the no-op before they call the codec.
    assert_eq!(requirements.allocations, 1);
}

#[test]
fn duplicate_noncanonical_nonfinite_and_shape_inputs_refuse_before_output() {
    let duplicate = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([numeric(1.0), numeric(2.0)])]],
        &[1, 2, 3],
    ));
    let error = prepare_chart_data_rewrite(&duplicate, &[vec![Some(3.0)]], options(&duplicate))
        .expect_err("duplicate numeric field");
    assert_eq!(
        error
            .decode_error()
            .and_then(super::DecodeError::duplicate_singular_field),
        Some("TSCH.GridValue.numeric_value")
    );

    let noncanonical_numeric = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([
            vec![0x89, 0x00],
            1.0f64.to_bits().to_le_bytes().to_vec(),
        ])]],
        &[4, 5],
    ));
    let error = prepare_chart_data_rewrite(
        &noncanonical_numeric,
        &[vec![Some(3.0)]],
        options(&noncanonical_numeric),
    )
    .expect_err("noncanonical numeric field");
    assert!(error.decode_error().is_some());

    let source = rich_source();
    let error = prepare_chart_data_rewrite(
        &source,
        &[vec![Some(f64::NAN), None], vec![Some(0.0), None]],
        options(&source),
    )
    .expect_err("non-finite request value");
    assert!(error.is_non_finite_numeric());
    let error = prepare_chart_data_rewrite(&source, &[vec![Some(1.0)]], options(&source))
        .expect_err("row-count mismatch");
    assert!(error.is_shape_mismatch());
    let error = prepare_chart_data_rewrite(
        &source,
        &[vec![Some(1.0), None], vec![Some(2.0)]],
        options(&source),
    )
    .expect_err("column-count mismatch");
    assert!(error.is_shape_mismatch());
    assert!(matches!(
        error,
        super::RewriteError::Shape {
            expected_rows: 2,
            actual_rows: 2,
            expected_columns: 2,
            actual_columns: 1,
        }
    ));
}

#[test]
fn unknown_groups_are_counted_and_copied_verbatim_at_every_chart_level() {
    let (source, root_group, grid_group, value_group, nested_root_group) = grouped_source();
    let replacement = [vec![Some(8.5)]];
    let prepared = prepare_chart_data_rewrite(&source, &replacement, options(&source))
        .expect("grouped chart prepares");

    let plain_source = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([date(45_125.5), numeric(17.25)])]],
        &[0x31, 0x41],
    ));
    let plain_prepared =
        prepare_chart_data_rewrite(&plain_source, &replacement, options(&plain_source))
            .expect("ungrouped chart prepares");
    assert!(
        prepared.execution_requirements().fields > plain_prepared.execution_requirements().fields,
        "nested group fields must be included in rewrite accounting"
    );

    let output = prepared
        .execute(prepared.execution_requirements().exact())
        .expect("grouped chart executes under exact requirements");
    assert!(output.report().changed());
    assert_ne!(output.output(), source.as_slice());
    for preserved_group in [&root_group, &nested_root_group, &grid_group, &value_group] {
        assert!(
            output
                .output()
                .windows(preserved_group.len())
                .any(|window| window == preserved_group.as_slice()),
            "candidate changed unknown group bytes: {preserved_group:?}"
        );
    }

    let snapshot = decode_modern(output.output(), &options(output.output())).expect("readback");
    assert_eq!(values(snapshot), [Some(8.5)]);
    assert_eq!(snapshot.row_labels().iter().collect::<Vec<_>>(), ["North"]);
    assert_eq!(snapshot.column_labels().iter().collect::<Vec<_>>(), ["Q1"]);
}

#[test]
fn mismatched_unknown_group_end_is_rejected_before_rewrite() {
    let mut malformed = group_tag(10, 3);
    malformed.extend(varint_field(4001, 0x55));
    malformed.extend(group_tag(11, 4));
    malformed.extend(modern(&grid(
        &["North"],
        &["Q1"],
        [vec![numeric(17.25)]],
        &[0x31, 0x41],
    )));

    let error = prepare_chart_data_rewrite(&malformed, &[vec![Some(8.5)]], options(&malformed))
        .expect_err("mismatched group terminator");
    assert!(error.decode_error().is_some());
}

#[test]
fn mixed_date_duration_values_keep_non_numeric_fields_when_cleared_or_added() {
    let source = modern(&grid(
        &["North"],
        &["Q1", "Q2"],
        [vec![
            joined([numeric(4.0), date(45_000.0), duration(2.0)]),
            joined([date(45_001.0), duration(3.0), length_field(77, b"opaque")]),
        ]],
        &[0xaa, 0xbb],
    ));
    let requested = [vec![None, Some(8.5)]];
    let prepared =
        prepare_chart_data_rewrite(&source, &requested, options(&source)).expect("prepare");
    let output = prepared
        .execute(prepared.execution_requirements().exact())
        .expect("execute");

    for preserved in [
        45_000.0f64.to_bits().to_le_bytes().as_slice(),
        2.0f64.to_bits().to_le_bytes().as_slice(),
        45_001.0f64.to_bits().to_le_bytes().as_slice(),
        3.0f64.to_bits().to_le_bytes().as_slice(),
        b"opaque".as_slice(),
        &[0xaa, 0xbb],
    ] {
        assert!(
            output
                .output()
                .windows(preserved.len())
                .any(|window| window == preserved)
        );
    }
    let snapshot = decode_modern(output.output(), &options(output.output())).expect("readback");
    assert_eq!(values(snapshot), [None, Some(8.5)]);
}

#[test]
fn standalone_numeric_addition_and_removal_resize_only_the_value_field() {
    let addition_source = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([date(45_000.0), length_field(77, b"opaque")])]],
        &[0x10, 0x20],
    ));
    let addition_values = [vec![Some(8.5)]];
    let prepared = prepare_chart_data_rewrite(
        &addition_source,
        &addition_values,
        options(&addition_source),
    )
    .expect("missing numeric field prepares for addition");
    let addition = prepared
        .execute(prepared.execution_requirements().exact())
        .expect("missing numeric field executes for addition");
    assert!(addition.output().len() > addition_source.len());
    let addition_golden = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([
            date(45_000.0),
            length_field(77, b"opaque"),
            numeric(8.5),
        ])]],
        &[0x10, 0x20],
    ));
    assert_eq!(addition.output(), addition_golden.as_slice());
    assert_eq!(
        values(decode_modern(addition.output(), &options(addition.output())).expect("readback")),
        [Some(8.5)]
    );
    assert!(
        addition
            .output()
            .windows(6)
            .any(|window| window == b"opaque")
    );

    let removal_source = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([
            numeric(8.5),
            date(45_000.0),
            length_field(77, b"opaque"),
        ])]],
        &[0x10, 0x20],
    ));
    let removal_values = [vec![None]];
    let prepared =
        prepare_chart_data_rewrite(&removal_source, &removal_values, options(&removal_source))
            .expect("numeric field prepares for removal");
    let removal = prepared
        .execute(prepared.execution_requirements().exact())
        .expect("numeric field executes for removal");
    assert!(removal.output().len() < removal_source.len());
    let removal_golden = modern(&grid(
        &["North"],
        &["Q1"],
        [vec![joined([date(45_000.0), length_field(77, b"opaque")])]],
        &[0x10, 0x20],
    ));
    assert_eq!(removal.output(), removal_golden.as_slice());
    assert_eq!(
        values(decode_modern(removal.output(), &options(removal.output())).expect("readback")),
        [None]
    );
    assert!(
        removal
            .output()
            .windows(6)
            .any(|window| window == b"opaque")
    );
}

#[test]
fn prepared_execution_accepts_exact_limits_and_rejects_one_below_each_preallocation_ceiling() {
    let source = rich_source();
    let replacement = [vec![Some(1.5), None], vec![Some(-0.0), Some(7.75)]];
    let prepared =
        prepare_chart_data_rewrite(&source, &replacement, options(&source)).expect("prepare");
    let requirements = prepared.execution_requirements();
    prepared
        .execute(requirements.exact())
        .expect("inclusive exact limits");

    let mut below = Vec::new();
    if requirements.output_bytes > 0 {
        below.push(
            requirements
                .exact()
                .with_output_bytes(requirements.output_bytes - 1),
        );
    }
    if requirements.fields > 0 {
        below.push(requirements.exact().with_fields(requirements.fields - 1));
    }
    if requirements.work_bytes > 0 {
        below.push(
            requirements
                .exact()
                .with_work_bytes(requirements.work_bytes - 1),
        );
    }
    if requirements.max_depth > 0 {
        below.push(
            requirements
                .exact()
                .with_max_depth(requirements.max_depth - 1),
        );
    }
    if requirements.allocations > 0 {
        below.push(
            requirements
                .exact()
                .with_allocations(requirements.allocations - 1),
        );
    }
    if requirements.retained_bytes > 0 {
        below.push(
            requirements
                .exact()
                .with_retained_bytes(requirements.retained_bytes - 1),
        );
    }
    if requirements.scratch_bytes > 0 {
        below.push(
            requirements
                .exact()
                .with_scratch_bytes(requirements.scratch_bytes - 1),
        );
    }

    for limits in below {
        let prepared = prepare_chart_data_rewrite(&source, &replacement, options(&source))
            .expect("prepare under broad limits");
        assert!(
            prepared.execute(limits).is_err(),
            "below exact limit was accepted"
        );
    }
}

#[test]
fn source_limits_are_checked_during_prepare_before_candidate_allocation() {
    let source = rich_source();
    let values = [vec![Some(1.5), None], vec![Some(-0.0), Some(7.75)]];
    let broad = options(&source);
    let prepared = prepare_chart_data_rewrite(&source, &values, broad).expect("baseline prepare");
    let requirements = prepared.execution_requirements();
    assert!(requirements.fields > 0);
    assert!(requirements.work_bytes > 0);

    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_input_bytes(source.len() - 1)
        )
        .is_err()
    );
    assert!(prepare_chart_data_rewrite(&source, &values, broad.with_max_fields(0)).is_err());
    assert!(prepare_chart_data_rewrite(&source, &values, broad.with_max_work_bytes(0)).is_err());
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_depth(requirements.max_depth - 1),
        )
        .is_err()
    );
    assert!(prepare_chart_data_rewrite(&source, &values, broad.with_max_output_bytes(0)).is_err());
    assert!(prepare_chart_data_rewrite(&source, &values, broad.with_max_allocations(0)).is_err());
    assert!(
        prepare_chart_data_rewrite(&source, &values, broad.with_max_retained_bytes(0)).is_err()
    );
    assert!(prepare_chart_data_rewrite(&source, &values, broad.with_max_scratch_bytes(0)).is_err());

    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_fields(requirements.fields - 1),
        )
        .is_err()
    );
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_work_bytes(requirements.work_bytes - 1),
        )
        .is_err()
    );
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_output_bytes(requirements.output_bytes - 1),
        )
        .is_err()
    );
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_allocations(requirements.allocations - 1),
        )
        .is_err()
    );
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_retained_bytes(requirements.retained_bytes - 1),
        )
        .is_err()
    );
    assert!(
        prepare_chart_data_rewrite(
            &source,
            &values,
            broad.with_max_scratch_bytes(requirements.scratch_bytes - 1),
        )
        .is_err()
    );
}
