use super::chart_grid_creation_codec::{
    ChartGridCreationRequest, EncodeError, EncodeLimit, EncodeOptions, InvalidInput,
    PreparedChartGridCreation, prepare_chart_grid_creation,
};

use prost::Message;

fn request<'a>(
    row_labels: &'a [String],
    column_labels: &'a [String],
    values: &'a [Vec<Option<f64>>],
    seed: u64,
) -> ChartGridCreationRequest<'a> {
    ChartGridCreationRequest::new(row_labels, column_labels, values, seed)
}

fn options<'a>(request: &ChartGridCreationRequest<'a>) -> EncodeOptions {
    EncodeOptions::for_request(request)
}

fn native_grid(
    row_labels: &[String],
    column_labels: &[String],
    values: &[Vec<Option<f64>>],
    seed: u64,
) -> crate::tsch::ChartGridArchive {
    let row_id_map = row_labels
        .iter()
        .enumerate()
        .map(
            |(index, _)| crate::tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                unique_id: deterministic_uuid(seed.wrapping_add(index as u64)),
                index: index as u32,
            },
        )
        .collect();
    let column_id_map = column_labels
        .iter()
        .enumerate()
        .map(
            |(index, _)| crate::tsch::chart_grid_archive::chart_grid_row_column_id_map::Entry {
                unique_id: deterministic_uuid(
                    seed.wrapping_add(1_u64 << 47).wrapping_add(index as u64),
                ),
                index: index as u32,
            },
        )
        .collect();

    crate::tsch::ChartGridArchive {
        row_name: row_labels.to_vec(),
        column_name: column_labels.to_vec(),
        grid_row: values
            .iter()
            .map(|row| crate::tsch::GridRow {
                value: row
                    .iter()
                    .map(|value| crate::tsch::GridValue {
                        numeric_value: *value,
                        ..Default::default()
                    })
                    .collect(),
            })
            .collect(),
        id_map: Some(crate::tsch::chart_grid_archive::ChartGridRowColumnIdMap {
            row_id_map,
            column_id_map,
        }),
    }
}

fn deterministic_uuid(seed: u64) -> String {
    let suffix = seed & 0x0000_ffff_ffff_ffff;
    format!("00000000-0000-4000-8000-{suffix:012X}")
}

fn prepared<'a>(
    row_labels: &'a [String],
    column_labels: &'a [String],
    values: &'a [Vec<Option<f64>>],
    seed: u64,
) -> PreparedChartGridCreation<'a> {
    let request = request(row_labels, column_labels, values, seed);
    let options = options(&request);
    prepare_chart_grid_creation(request, options).expect("valid chart grid request")
}

#[test]
fn creation_matches_the_native_prost_grid_byte_for_byte() {
    let row_labels = vec![String::from("North"), String::from("South")];
    let column_labels = vec![String::from("Q1"), String::from("Q2")];
    let values = vec![vec![Some(17.25), None], vec![Some(-8.5), Some(42.0)]];
    let request = request(&row_labels, &column_labels, &values, 0x0102_0304_0506_0708);
    let options = options(&request);
    let output = super::chart_grid_creation_codec::encode_chart_grid(request, options)
        .expect("encode chart grid");
    let expected =
        native_grid(&row_labels, &column_labels, &values, 0x0102_0304_0506_0708).encode_to_vec();

    let encoded = output.into_bytes();
    assert_eq!(encoded, expected);
    let decoded = crate::tsch::ChartGridArchive::decode(encoded.as_slice()).expect("prost decode");
    assert_eq!(
        decoded,
        native_grid(&row_labels, &column_labels, &values, 0x0102_0304_0506_0708,)
    );
}

#[test]
fn deterministic_ids_match_the_legacy_builder_across_wrapping_seeds() {
    let row_labels = vec![String::from("r0"), String::from("r1")];
    let column_labels = vec![String::from("c0"), String::from("c1")];
    let values = vec![vec![Some(1.0), Some(2.0)], vec![None, Some(3.0)]];

    for seed in [0, 1, (1_u64 << 47) - 1, 1_u64 << 47, u64::MAX - 1, u64::MAX] {
        let request = request(&row_labels, &column_labels, &values, seed);
        let options = options(&request);
        let output = super::chart_grid_creation_codec::encode_chart_grid(request, options)
            .expect("edge seed is valid");
        let encoded = output.into_bytes();
        assert_eq!(
            encoded,
            native_grid(&row_labels, &column_labels, &values, seed).encode_to_vec()
        );
    }
}

#[test]
fn signed_zero_and_missing_numeric_cells_keep_presence_and_bits() {
    let row_labels = vec![String::from("row")];
    let column_labels = vec![
        String::from("negative"),
        String::from("empty"),
        String::from("positive"),
    ];
    let values = vec![vec![Some(-0.0), None, Some(0.0)]];
    let request = request(&row_labels, &column_labels, &values, 7);
    let options = options(&request);
    let bytes = super::chart_grid_creation_codec::encode_chart_grid(request, options)
        .expect("signed zero grid")
        .into_bytes();
    let decoded = crate::tsch::ChartGridArchive::decode(bytes.as_slice()).expect("prost decode");
    let cells = &decoded.grid_row[0].value;
    assert_eq!(
        cells[0].numeric_value.expect("negative zero").to_bits(),
        (-0.0_f64).to_bits()
    );
    assert_eq!(cells[1].numeric_value, None);
    assert_eq!(
        cells[2].numeric_value.expect("positive zero").to_bits(),
        0.0_f64.to_bits()
    );
}

#[test]
fn unicode_labels_are_encoded_as_valid_utf8_without_normalization() {
    let row_labels = vec![String::from("北至"), String::from("إيرادات 🧪")];
    let column_labels = vec![String::from("第一季度"), String::from("résumé")];
    let values = vec![vec![Some(1.0), None], vec![None, Some(2.0)]];
    let request = request(&row_labels, &column_labels, &values, 19);
    let options = options(&request);
    let output = super::chart_grid_creation_codec::encode_chart_grid(request, options)
        .expect("unicode labels");
    let encoded = output.into_bytes();
    let decoded = crate::tsch::ChartGridArchive::decode(encoded.as_slice()).expect("prost decode");
    assert_eq!(decoded.row_name, row_labels);
    assert_eq!(decoded.column_name, column_labels);
}

#[test]
fn empty_axes_and_ragged_values_are_rejected_before_encoding() {
    let columns = vec![String::from("Q1")];
    let rows = vec![String::from("North"), String::from("South")];
    let valid_values = vec![vec![Some(1.0)], vec![Some(2.0)]];

    let cases = [
        (Vec::new(), columns.clone(), valid_values.clone()),
        (rows.clone(), Vec::new(), valid_values.clone()),
        (rows.clone(), columns.clone(), vec![vec![Some(1.0)]]),
        (rows.clone(), columns.clone(), vec![vec![Some(1.0)], vec![]]),
    ];
    for (row_labels, column_labels, values) in cases {
        let request = request(&row_labels, &column_labels, &values, 1);
        let options = options(&request);
        assert!(matches!(
            prepare_chart_grid_creation(request, options),
            Err(EncodeError::InvalidInput(_))
        ));
    }
}

#[test]
fn non_finite_values_are_rejected_with_coordinates() {
    let rows = vec![String::from("North")];
    let columns = vec![String::from("Q1"), String::from("Q2")];
    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let values = vec![vec![Some(1.0), Some(value)]];
        let request = request(&rows, &columns, &values, 1);
        let options = options(&request);
        let error = prepare_chart_grid_creation(request, options)
            .expect_err("non-finite values must be rejected");
        assert!(matches!(
            error,
            EncodeError::InvalidInput(InvalidInput::NonFiniteNumeric { row: 0, column: 1 })
        ));
    }
}

#[test]
fn prepared_execution_admits_every_exact_budget_and_rejects_one_below() {
    let rows = vec![String::from("North"), String::from("South")];
    let columns = vec![String::from("Q1"), String::from("Q2")];
    let values = vec![vec![Some(1.0), None], vec![Some(2.0), Some(3.0)]];
    let prepared = prepared(&rows, &columns, &values, 11);
    let requirements = prepared.execution_requirements();
    assert!(requirements.output_bytes > 0);
    assert!(requirements.fields > 0);
    assert!(requirements.work_bytes > 0);
    assert_eq!(requirements.max_depth, 3);
    assert!(requirements.allocations > 0);
    assert!(requirements.retained_bytes > 0);
    assert!(requirements.scratch_bytes > 0);
    assert!(requirements.cells > 0);
    assert!(requirements.labels > 0);
    assert!(requirements.text_bytes > 0);

    let output = prepared
        .execute(requirements.exact())
        .expect("exact requirements are inclusive");
    assert_eq!(output.report(), requirements);
    assert_eq!(output.bytes().len(), requirements.output_bytes);

    let below = [
        requirements
            .exact()
            .with_output_bytes(requirements.output_bytes - 1),
        requirements.exact().with_fields(requirements.fields - 1),
        requirements
            .exact()
            .with_work_bytes(requirements.work_bytes - 1),
        requirements
            .exact()
            .with_allocations(requirements.allocations - 1),
        requirements
            .exact()
            .with_retained_bytes(requirements.retained_bytes - 1),
        requirements
            .exact()
            .with_scratch_bytes(requirements.scratch_bytes - 1),
        requirements.exact().with_cells(requirements.cells - 1),
        requirements.exact().with_labels(requirements.labels - 1),
        requirements
            .exact()
            .with_text_bytes(requirements.text_bytes - 1),
        requirements
            .exact()
            .with_max_depth(requirements.max_depth - 1),
    ];
    for limits in below {
        assert!(
            prepared.execute(limits).is_err(),
            "one below an exact creation budget must be refused",
        );
    }
}

#[test]
fn zero_resource_limits_precede_late_nonfinite_validation() {
    let rows = vec![String::from("row")];
    let columns = vec![String::from("first"), String::from("last")];
    // Keep the invalid value last so a planner that scans semantic cells
    // before checking aggregate ceilings would accidentally report the
    // scalar error instead of the caller's selected resource limit.
    let values = vec![vec![Some(1.0), Some(f64::NAN)]];
    let request = request(&rows, &columns, &values, 23);
    let broad = options(&request);

    for (options, expected) in [
        (
            broad.with_max_cells(0),
            EncodeLimit::Cells {
                observed: 0,
                maximum: 0,
            },
        ),
        (
            broad.with_max_labels(0),
            EncodeLimit::Labels {
                observed: 0,
                maximum: 0,
            },
        ),
        (
            broad.with_max_work_bytes(0),
            EncodeLimit::Work {
                observed: 0,
                maximum: 0,
            },
        ),
    ] {
        let error = prepare_chart_grid_creation(request, options)
            .expect_err("the selected resource ceiling must be checked first");
        let Some(actual) = error.resource_limit() else {
            panic!("expected a typed resource limit, got {error:?}");
        };
        assert_eq!(
            std::mem::discriminant(&actual),
            std::mem::discriminant(&expected),
            "unexpected precedence error: {error:?}"
        );
        match (actual, expected) {
            (
                EncodeLimit::Cells { maximum, .. },
                EncodeLimit::Cells {
                    maximum: expected_maximum,
                    ..
                },
            )
            | (
                EncodeLimit::Labels { maximum, .. },
                EncodeLimit::Labels {
                    maximum: expected_maximum,
                    ..
                },
            )
            | (
                EncodeLimit::Work { maximum, .. },
                EncodeLimit::Work {
                    maximum: expected_maximum,
                    ..
                },
            ) => assert_eq!(maximum, expected_maximum),
            _ => panic!("unexpected resource limit: {actual:?}"),
        }
    }
}

#[test]
fn a_looser_nesting_ceiling_is_accepted_without_changing_the_wire_depth() {
    let rows = vec![String::from("row")];
    let columns = vec![String::from("column")];
    let values = vec![vec![Some(4.0)]];
    let request = request(&rows, &columns, &values, 29);
    let options = options(&request).with_max_depth(4);
    let prepared = prepare_chart_grid_creation(request, options)
        .expect("a caller ceiling looser than the fixed graph depth is valid");
    assert_eq!(prepared.execution_requirements().max_depth, 3);
    assert!(
        prepared
            .execute(prepared.execution_requirements().exact())
            .is_ok()
    );
}

#[test]
fn one_output_allocation_handles_many_scalar_buffa_leaves() {
    let rows = (0..4).map(|index| format!("r{index}")).collect::<Vec<_>>();
    let columns = (0..7).map(|index| format!("c{index}")).collect::<Vec<_>>();
    let values = (0..rows.len())
        .map(|row| {
            (0..columns.len())
                .map(|column| Some((row * columns.len() + column) as f64 + 0.25))
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    let request = request(&rows, &columns, &values, 31);
    let options = options(&request).with_max_allocations(1);
    let output = super::chart_grid_creation_codec::encode_chart_grid(request, options)
        .expect("one output allocation is enough");
    assert_eq!(output.report().allocations, 1);
    assert_eq!(
        output.bytes(),
        native_grid(&rows, &columns, &values, 31).encode_to_vec()
    );
}
