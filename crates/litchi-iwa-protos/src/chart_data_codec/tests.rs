use std::ptr;

use prost::Message as _;

use super::{
    ChartDataSnapshot, DecodeError, DecodeLimit, DecodeOptions, decode_grid,
    decode_grid_with_report, decode_modern, decode_modern_with_report,
};

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

fn joined(fields: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    fields.into_iter().flatten().collect()
}

fn numeric(value: f64) -> Vec<u8> {
    fixed64_field(1, value)
}

fn date_value(value: f64) -> Vec<u8> {
    fixed64_field(4, value)
}

fn duration_value(value: f64) -> Vec<u8> {
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
    id_map: Option<&[u8]>,
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
    if let Some(id_map) = id_map {
        output.extend(length_field(4, id_map));
    }
    output
}

fn modern(grid: &[u8]) -> Vec<u8> {
    let chart = joined([
        length_field(7, grid),
        // Unknown chart fields are intentionally left to the source path.
        varint_field(91, 0xfeed),
    ]);
    joined([
        length_field(10_000, &chart),
        // Unknown drawable fields must not affect the selected extension.
        length_field(92, b"future drawable data"),
    ])
}

fn permissive(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len(),
        usize::MAX,
        usize::MAX,
        64,
        usize::MAX,
        usize::MAX,
        usize::MAX,
    )
}

fn embedded_subslice<'source>(source: &'source [u8], needle: &[u8]) -> &'source [u8] {
    let offset = source
        .windows(needle.len())
        .position(|window| window == needle)
        .expect("fixture contains embedded payload");
    &source[offset..offset + needle.len()]
}

fn rich_grid() -> Vec<u8> {
    grid(
        &["North", "South"],
        &["Q1", "Q2"],
        [
            vec![numeric(0.0), date_value(45_123.25)],
            vec![duration_value(3.5), numeric(-9.25)],
        ],
        Some(&[0xff, 0x80, 0x80, 0x01]),
    )
}

fn rich_modern() -> Vec<u8> {
    modern(&rich_grid())
}

fn assert_limit(error: DecodeError, expected: fn(DecodeLimit) -> bool) {
    let limit = error
        .resource_limit()
        .expect("failure should identify its bounded resource");
    assert!(expected(limit), "unexpected resource limit: {limit:?}");
    assert!(error.report().source_bytes() > 0);
}

fn all_values(snapshot: ChartDataSnapshot<'_>) -> Vec<Option<f64>> {
    snapshot
        .rows()
        .iter()
        .flat_map(|row| row.values())
        .collect()
}

#[test]
fn modern_data_keeps_optional_numeric_presence_and_borrows_every_repetition() {
    let grid_source = rich_grid();
    let source = modern(&grid_source);
    let options = permissive(&source);
    let (snapshot, report) = decode_modern_with_report(&source, &options).expect("modern chart");

    assert!(ptr::eq(snapshot.source(), source.as_slice()));
    assert!(ptr::eq(
        snapshot.grid_source(),
        embedded_subslice(&source, &grid_source)
    ));
    assert_eq!(
        snapshot.row_labels().iter().collect::<Vec<_>>(),
        ["North", "South"]
    );
    assert_eq!(
        snapshot.column_labels().iter().collect::<Vec<_>>(),
        ["Q1", "Q2"]
    );
    assert_eq!(snapshot.row_count(), 2);
    assert_eq!(snapshot.column_count(), 2);
    assert_eq!(snapshot.rows().len(), 2);
    assert_eq!(
        snapshot
            .rows()
            .iter()
            .map(|row| row.len())
            .collect::<Vec<_>>(),
        [2, 2]
    );
    assert_eq!(all_values(snapshot), [Some(0.0), None, None, Some(-9.25)]);
    assert_eq!(report.source_bytes(), source.len());
    assert_eq!(report.cell_count(), 4);
    assert_eq!(report.label_count(), 4);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), source.len());

    // Explicit zero is present, while a date/duration-only GridValue has no
    // numeric value in the semantic projection.
    assert_eq!(all_values(snapshot)[0].expect("explicit zero").to_bits(), 0);
    assert!(all_values(snapshot)[1].is_none());
}

#[test]
fn direct_grid_matches_the_native_prost_wire_oracle() {
    let native = crate::tsch::ChartGridArchive {
        row_name: vec![String::from("North"), String::from("South")],
        column_name: vec![String::from("Q1"), String::from("Q2")],
        grid_row: vec![
            crate::tsch::GridRow {
                value: vec![
                    crate::tsch::GridValue {
                        numeric_value: Some(0.0),
                        ..Default::default()
                    },
                    crate::tsch::GridValue {
                        date_value: Some(45_123.25),
                        ..Default::default()
                    },
                ],
            },
            crate::tsch::GridRow {
                value: vec![
                    crate::tsch::GridValue {
                        duration_value: Some(3.5),
                        ..Default::default()
                    },
                    crate::tsch::GridValue {
                        numeric_value: Some(-9.25),
                        ..Default::default()
                    },
                ],
            },
        ],
        id_map: None,
    };
    let source = native.encode_to_vec();
    let decoded_native = crate::tsch::ChartGridArchive::decode(source.as_slice())
        .expect("prost should decode its own canonical fixture");
    assert_eq!(decoded_native, native);

    let snapshot = decode_grid(&source, &permissive(&source)).expect("grid data");
    assert_eq!(
        snapshot.row_labels().iter().collect::<Vec<_>>(),
        ["North", "South"]
    );
    assert_eq!(
        snapshot.column_labels().iter().collect::<Vec<_>>(),
        ["Q1", "Q2"]
    );
    assert_eq!(all_values(snapshot), [Some(0.0), None, None, Some(-9.25)]);
}

#[test]
fn opaque_id_map_unknown_fields_and_date_semantics_are_ignored() {
    let grid_source = grid(
        &["North"],
        &["Q1"],
        [vec![joined([
            // Numeric value remains authoritative when date fields coexist.
            fixed64_field(1, 17.25),
            date_value(45123.0),
            duration_value(0.5),
            // Unknown fields are accepted but never materialized.
            varint_field(99, 7),
        ])]],
        Some(&[0xff, 0xff, 0xff]),
    );
    let mut source = grid_source.clone();
    source.extend(length_field(100, b"future grid extension"));
    let (snapshot, report) = decode_grid_with_report(&source, &permissive(&source))
        .expect("opaque future fields must not affect numeric projection");

    assert_eq!(snapshot.row_count(), 1);
    assert_eq!(snapshot.column_count(), 1);
    assert_eq!(all_values(snapshot), [Some(17.25)]);
    assert_eq!(report.cell_count(), 1);
    assert_eq!(report.allocations(), 0);
    assert!(ptr::eq(snapshot.grid_source(), source.as_slice()));
}

#[test]
fn empty_and_ragged_grids_fail_with_typed_shape_errors() {
    let empty_cases = [
        grid(&[], &["Q1"], Vec::<Vec<Vec<u8>>>::new(), None),
        grid(&["North"], &[], Vec::<Vec<Vec<u8>>>::new(), None),
        grid(&["North"], &["Q1"], Vec::<Vec<Vec<u8>>>::new(), None),
    ];
    for source in empty_cases {
        let error = decode_grid(&source, &permissive(&source)).expect_err("empty grid");
        assert!(error.is_non_rectangular(), "{error}");
    }

    let ragged = grid(
        &["North", "South"],
        &["Q1", "Q2"],
        [vec![numeric(1.0), numeric(2.0)], vec![numeric(3.0)]],
        None,
    );
    let error = decode_grid(&ragged, &permissive(&ragged)).expect_err("ragged row");
    assert!(error.is_non_rectangular(), "{error}");

    let wrong_row_count = grid(
        &["North"],
        &["Q1"],
        [vec![numeric(1.0)], vec![numeric(2.0)]],
        None,
    );
    let error = decode_grid(&wrong_row_count, &permissive(&wrong_row_count))
        .expect_err("row-label mismatch");
    assert!(error.is_non_rectangular(), "{error}");
}

#[test]
fn duplicate_selected_fields_are_rejected_even_when_the_first_value_is_zero() {
    let duplicate_numeric = grid(
        &["North"],
        &["Q1"],
        [vec![joined([fixed64_field(1, 0.0), fixed64_field(1, 4.0)])]],
        None,
    );
    let error = decode_grid(&duplicate_numeric, &permissive(&duplicate_numeric))
        .expect_err("duplicate numeric field");
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.GridValue.numeric_value")
    );

    let duplicate_date = grid(
        &["North"],
        &["Q1"],
        [vec![joined([date_value(0.0), date_value(1.0)])]],
        None,
    );
    let error = decode_grid(&duplicate_date, &permissive(&duplicate_date))
        .expect_err("duplicate date field");
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.GridValue.date_value")
    );

    let valid_grid = grid(&["North"], &["Q1"], [vec![numeric(1.0)]], None);
    let duplicate_grid = joined([length_field(7, &valid_grid), length_field(7, &valid_grid)]);
    let duplicate_chart = length_field(10_000, &duplicate_grid);
    let error = decode_modern(&duplicate_chart, &permissive(&duplicate_chart))
        .expect_err("duplicate ChartArchive.grid");
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartArchive.grid")
    );

    let duplicate_extension = joined([
        length_field(10_000, &length_field(7, &valid_grid)),
        length_field(10_000, &length_field(7, &valid_grid)),
    ]);
    let error = decode_modern(&duplicate_extension, &permissive(&duplicate_extension))
        .expect_err("duplicate unity extension");
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartDrawableArchive.unity")
    );
}

#[test]
fn wrong_wire_types_and_nonfinite_numeric_values_fail_closed() {
    let wrong_grid_wire = varint_field(7, 1);
    let chart = length_field(10_000, &wrong_grid_wire);
    assert!(decode_modern(&chart, &permissive(&chart)).is_err());

    let wrong_row_wire = joined([
        length_field(1, b"North"),
        length_field(2, b"Q1"),
        fixed64_field(3, 1.0),
    ]);
    assert!(decode_grid(&wrong_row_wire, &permissive(&wrong_row_wire)).is_err());

    let wrong_value_wire = grid(&["North"], &["Q1"], [vec![varint_field(1, 1)]], None);
    assert!(decode_grid(&wrong_value_wire, &permissive(&wrong_value_wire)).is_err());

    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        let source = grid(&["North"], &["Q1"], [vec![numeric(value)]], None);
        let error = decode_grid(&source, &permissive(&source))
            .expect_err("non-finite numeric values are not chart data");
        assert!(error.is_non_finite_numeric(), "{error}");
    }
}

#[test]
fn modern_decoder_requires_the_selected_extension_and_grid() {
    let missing_extension = length_field(99, b"future");
    let error = decode_modern(&missing_extension, &permissive(&missing_extension))
        .expect_err("missing unity extension");
    assert_eq!(
        error.missing_required_field(),
        Some("TSCH.ChartDrawableArchive.unity")
    );

    let missing_grid = length_field(10_000, &varint_field(1, 2));
    let error =
        decode_modern(&missing_grid, &permissive(&missing_grid)).expect_err("missing chart grid");
    assert_eq!(
        error.missing_required_field(),
        Some("TSCH.ChartArchive.grid")
    );
}

#[test]
fn every_decode_budget_is_inclusive_and_failure_reports_are_complete() {
    let source = rich_modern();
    let broad = permissive(&source);
    let (_, report) = decode_modern_with_report(&source, &broad).expect("baseline");
    assert_eq!(report.source_bytes(), source.len());

    let exact = DecodeOptions::new(
        report.source_bytes(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
        report.cell_count(),
        report.label_count(),
        report.text_bytes(),
    );
    decode_modern(&source, &exact).expect("all observed budgets are inclusive");

    let error = decode_modern(
        &source,
        &broad.with_max_input_bytes(report.source_bytes() - 1),
    )
    .expect_err("one byte below input budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Bytes { .. }));

    let error = decode_modern(&source, &broad.with_max_fields(report.fields() - 1))
        .expect_err("one field below field budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Fields { .. }));

    let error = decode_modern(&source, &broad.with_max_work_bytes(report.work_bytes() - 1))
        .expect_err("one byte below work budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Work { .. }));

    let error = decode_modern(&source, &broad.with_max_depth(report.max_depth() - 1))
        .expect_err("one level below nesting budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Nesting { .. }));

    let error = decode_modern(&source, &broad.with_max_cells(report.cell_count() - 1))
        .expect_err("one cell below cell budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Cells { .. }));

    let error = decode_modern(
        &source,
        &broad.with_max_label_count(report.label_count() - 1),
    )
    .expect_err("one label below label budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Labels { .. }));

    let error = decode_modern(&source, &broad.with_max_text_bytes(report.text_bytes() - 1))
        .expect_err("one text byte below text budget");
    assert_limit(error, |limit| matches!(limit, DecodeLimit::Text { .. }));

    // A failed decode still reports its source size and the work charged up
    // to the rejected operation; callers can make an informed retry policy.
    assert!(error_report_has_source_and_work(
        decode_modern(&source, &broad.with_max_work_bytes(report.work_bytes() - 1))
            .expect_err("work budget")
            .report()
    ));
}

#[test]
fn direct_grid_input_budget_is_inclusive_and_precedes_work() {
    let source = rich_grid();
    let exact = permissive(&source).with_max_input_bytes(source.len());
    decode_grid(&source, &exact).expect("an exact direct-grid input limit is inclusive");

    let hostile = permissive(&source)
        .with_max_input_bytes(source.len() - 1)
        .with_max_work_bytes(0);
    let error = decode_grid(&source, &hostile)
        .expect_err("the direct grid must reject oversized input before traversal work");
    assert_eq!(
        error.resource_limit(),
        Some(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: source.len() - 1,
        })
    );
    assert_eq!(error.report().source_bytes(), source.len());
    assert_eq!(error.report().work_bytes(), 0);
}

fn error_report_has_source_and_work(report: super::DecodeReport) -> bool {
    report.source_bytes() > 0 && report.work_bytes() > 0 && report.retained_bytes() > 0
}
