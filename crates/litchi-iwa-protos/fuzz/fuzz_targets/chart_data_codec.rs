#![no_main]

use std::{hint::black_box, ptr, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::chart_data_codec::{
    ChartDataSnapshot, DecodeError, DecodeLimit, DecodeOptions, RewriteError,
    RewriteExecutionRequirements, decode_grid, decode_grid_with_report, decode_modern,
    decode_modern_with_report, prepare_chart_data_rewrite,
};

// Keep both the fuzzer input and every decoder pass bounded. Oversized inputs
// are skipped rather than truncated, so all routes receive one unchanged
// caller-owned source.
const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_DEPTH: u32 = 64;
const MAX_CELLS: usize = 8 * 1024;
const MAX_LABELS: usize = 8 * 1024;
const MAX_TEXT_BYTES: usize = 64 * 1024;

#[derive(Clone, Copy)]
struct SourceRange {
    start: usize,
    end: usize,
}

impl SourceRange {
    fn new(source: &[u8]) -> Self {
        let start = source.as_ptr() as usize;
        let end = start
            .checked_add(source.len())
            .expect("bounded source pointer range");
        Self { start, end }
    }

    fn assert_borrowed(self, bytes: &[u8]) {
        if bytes.is_empty() {
            return;
        }
        let start = bytes.as_ptr() as usize;
        let end = start
            .checked_add(bytes.len())
            .expect("bounded borrowed payload range");
        assert!(
            start >= self.start && end <= self.end,
            "decoded payload did not borrow from the caller source"
        );
    }
}

#[derive(Clone, Copy)]
enum Route {
    Modern,
    Grid,
}

impl Route {
    fn decode<'source>(
        self,
        source: &'source [u8],
        options: &DecodeOptions,
    ) -> Result<ChartDataSnapshot<'source>, DecodeError> {
        match self {
            Self::Modern => decode_modern(source, options),
            Self::Grid => decode_grid(source, options),
        }
    }

    fn decode_with_report<'source>(
        self,
        source: &'source [u8],
        options: &DecodeOptions,
    ) -> Result<
        (
            ChartDataSnapshot<'source>,
            litchi_iwa_protos::chart_data_codec::DecodeReport,
        ),
        DecodeError,
    > {
        match self {
            Self::Modern => decode_modern_with_report(source, options),
            Self::Grid => decode_grid_with_report(source, options),
        }
    }
}

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    // Arbitrary data drives both routes. A fixed one-time matrix keeps valid,
    // malformed, non-finite, ragged, and duplicate-field cases hot even when
    // a mutation campaign has not yet discovered a complete chart grid.
    exercise_route(&source, Route::Modern, data);
    exercise_route(&source, Route::Grid, data);

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for (source, route) in fixed_cases() {
            exercise_route(&source, route, b"fixed-chart-data-case");
        }
    });
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn options() -> DecodeOptions {
    DecodeOptions::new(
        MAX_INPUT_BYTES,
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_DEPTH,
        MAX_CELLS,
        MAX_LABELS,
        MAX_TEXT_BYTES,
    )
}

fn exercise_route(source: &[u8], route: Route, request_seed: &[u8]) {
    let before = source.to_vec();
    let source_range = SourceRange::new(source);
    let limits = options();
    let scalar = route.decode(source, &limits);
    assert_eq!(
        source,
        before.as_slice(),
        "scalar decode modified its source"
    );

    let reported = route.decode_with_report(source, &limits);
    assert_eq!(
        source,
        before.as_slice(),
        "reported decode modified its source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            observe_snapshot(snapshot, source_range, route);
            assert_report(report, source.len());
            black_box((snapshot, report));

            // A successful source is also replayed with its exact observed
            // accounting. This exercises inclusive limits and the complete
            // failure report for each bounded resource axis.
            exercise_limit_failures(source, route, report);

            // The source-preserving writer only accepts a complete modern
            // drawable.  Direct-grid decoding remains read-only, so keeping
            // this branch explicit prevents a writer from accidentally
            // treating a grid payload as a drawable root.
            if matches!(route, Route::Modern) {
                exercise_rewrites(source, snapshot, request_seed, &before);
            }
        },
        (Err(error), Err(reported_error)) => {
            observe_error(error);
            observe_error(reported_error);
        },
        (scalar, reported) => {
            panic!(
                "scalar/report decode disagreement: scalar={:?}, reported={:?}",
                scalar.map(|snapshot| snapshot.row_count()),
                reported.map(|(snapshot, report)| (snapshot.row_count(), report)),
            );
        },
    }
}

fn snapshot_values(snapshot: ChartDataSnapshot<'_>) -> Vec<Vec<Option<f64>>> {
    snapshot
        .rows()
        .iter()
        .map(|row| row.values().collect())
        .collect()
}

fn generated_value(data: &[u8], cell: usize) -> f64 {
    let first = data
        .get(cell % data.len().max(1))
        .copied()
        .unwrap_or_default();
    let second = data
        .get((cell.wrapping_mul(7).wrapping_add(3)) % data.len().max(1))
        .copied()
        .unwrap_or_default();
    let magnitude = f64::from(first) + f64::from(second) / 256.0 + 0.25;
    if first & 1 == 0 {
        magnitude
    } else {
        -magnitude
    }
}

fn replacement_values(snapshot: ChartDataSnapshot<'_>, data: &[u8]) -> Vec<Vec<Option<f64>>> {
    let mut cell = 0usize;
    snapshot
        .rows()
        .iter()
        .map(|row| {
            row.values()
                .map(|current| {
                    let marker = data
                        .get(cell % data.len().max(1))
                        .copied()
                        .unwrap_or_default();
                    let replacement = if marker % 3 == 0 {
                        None
                    } else {
                        Some(generated_value(data, cell))
                    };
                    cell = cell.checked_add(1).expect("bounded fuzz cell count");
                    // Keep the generated value finite while making both
                    // numeric insertion and numeric removal reachable.
                    let _ = current;
                    replacement
                })
                .collect()
        })
        .collect()
}

fn exercise_rewrites(source: &[u8], snapshot: ChartDataSnapshot<'_>, data: &[u8], before: &[u8]) {
    let original_values = snapshot_values(snapshot);
    let desired_values = replacement_values(snapshot, data);

    // A matching request must be a true byte-exact no-op, including unknown
    // outer/chart/grid fields and the opaque native id map.
    exercise_one_rewrite(source, &original_values, before, true);

    // A second request changes only numeric presence/bits while keeping the
    // validated row/column shape.  The writer's readback is compared bitwise
    // so signed zero remains observable.
    exercise_one_rewrite(
        source,
        &desired_values,
        before,
        value_grids_equal(&original_values, &desired_values),
    );
}

fn value_grids_equal(left: &[Vec<Option<f64>>], right: &[Vec<Option<f64>>]) -> bool {
    left.len() == right.len()
        && left.iter().zip(right).all(|(left, right)| {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| match (left, right) {
                        (None, None) => true,
                        (Some(left), Some(right)) => left.to_bits() == right.to_bits(),
                        _ => false,
                    })
        })
}

fn exercise_one_rewrite(
    source: &[u8],
    desired: &[Vec<Option<f64>>],
    before: &[u8],
    expected_noop: bool,
) {
    let options = options();
    let prepared = prepare_chart_data_rewrite(source, desired, options);
    let Ok(prepared) = prepared else {
        // Random valid modern payloads can legitimately hit the finite output
        // ceiling when a missing numeric value is added.  Preparation itself
        // is the bounded refusal path; fixed cases cover successful writes.
        return;
    };
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(requirements.exact())
        .unwrap_or_else(|error| panic!("exact chart-data rewrite failed: {error:?}"));
    assert_eq!(source, before, "chart-data rewrite modified its source");
    assert_eq!(output.report().output_bytes(), output.output().len());
    assert_eq!(output.report().changed(), output.output() != source);
    assert_eq!(output.output() == source, expected_noop);
    if expected_noop {
        assert_eq!(requirements.output_bytes, source.len());
        // The package-level snapshot is allocation-free, while the codec
        // returns an owned Vec even for a byte-exact no-op.
        assert_eq!(requirements.allocations, usize::from(!source.is_empty()));
    } else {
        assert_preserved_wire_markers(source, output.output());
    }

    let decoded = decode_modern(output.output(), &options)
        .unwrap_or_else(|error| panic!("chart-data rewrite readback failed: {error:?}"));
    assert_eq!(snapshot_values(decoded), desired);
    exercise_rewrite_limit_failures(source, desired, requirements);
    black_box(output);
}

fn assert_preserved_wire_markers(source: &[u8], output: &[u8]) {
    // These markers occur in the fixed rich source and make preservation of
    // unknown fields observable without parsing private wire spans here.
    for marker in [
        b"future drawable extension".as_slice(),
        b"future grid extension".as_slice(),
        b"future chart extension".as_slice(),
        b"cell future bytes".as_slice(),
        b"blank future bytes".as_slice(),
    ] {
        if source.windows(marker.len()).any(|window| window == marker) {
            assert!(
                output.windows(marker.len()).any(|window| window == marker),
                "chart-data rewrite dropped preserved wire marker"
            );
        }
    }
}

fn exercise_rewrite_limit_failures(
    source: &[u8],
    desired: &[Vec<Option<f64>>],
    requirements: RewriteExecutionRequirements,
) {
    let probes = [
        (requirements.output_bytes > 0).then(|| {
            requirements
                .exact()
                .with_output_bytes(requirements.output_bytes - 1)
        }),
        (requirements.fields > 0)
            .then(|| requirements.exact().with_fields(requirements.fields - 1)),
        (requirements.work_bytes > 0).then(|| {
            requirements
                .exact()
                .with_work_bytes(requirements.work_bytes - 1)
        }),
        (requirements.max_depth > 0).then(|| {
            requirements
                .exact()
                .with_max_depth(requirements.max_depth - 1)
        }),
        (requirements.allocations > 0).then(|| {
            requirements
                .exact()
                .with_allocations(requirements.allocations - 1)
        }),
        (requirements.retained_bytes > 0).then(|| {
            requirements
                .exact()
                .with_retained_bytes(requirements.retained_bytes - 1)
        }),
        (requirements.scratch_bytes > 0).then(|| {
            requirements
                .exact()
                .with_scratch_bytes(requirements.scratch_bytes - 1)
        }),
    ];
    for limits in probes.into_iter().flatten() {
        let prepared = prepare_chart_data_rewrite(source, desired, options())
            .expect("broad chart-data rewrite limits must prepare");
        let error = prepared
            .execute(limits)
            .expect_err("one-below chart-data rewrite limit was accepted");
        assert!(error.resource_limit().is_some());
        observe_rewrite_error(error);
    }
}

fn observe_rewrite_error(error: RewriteError) {
    black_box((error.resource_limit(), error.allocation_amount()));
    black_box(format_args!("{error:?}"));
}

fn observe_snapshot(snapshot: ChartDataSnapshot<'_>, source: SourceRange, route: Route) {
    source.assert_borrowed(snapshot.source());
    source.assert_borrowed(snapshot.grid_source());
    if matches!(route, Route::Grid) {
        assert!(ptr::eq(snapshot.grid_source(), snapshot.source()));
    }

    let row_labels = snapshot.row_labels();
    let column_labels = snapshot.column_labels();
    assert_eq!(row_labels.len(), snapshot.row_count());
    assert_eq!(column_labels.len(), snapshot.column_count());
    for label in row_labels.iter().chain(column_labels.iter()) {
        source.assert_borrowed(label.as_bytes());
    }

    let mut rows = 0usize;
    for row in snapshot.rows().iter() {
        rows = rows.checked_add(1).expect("bounded row count");
        assert_eq!(row.len(), snapshot.column_count());
        let mut values = 0usize;
        for value in row.values() {
            values = values.checked_add(1).expect("bounded cell count");
            if let Some(value) = value {
                assert!(
                    value.is_finite(),
                    "successful decode exposed non-finite data"
                );
            }
        }
        assert_eq!(values, row.len());
    }
    assert_eq!(rows, snapshot.row_count());
}

fn assert_report(report: litchi_iwa_protos::chart_data_codec::DecodeReport, source_bytes: usize) {
    assert_eq!(report.source_bytes(), source_bytes);
    assert!(report.source_bytes() <= MAX_INPUT_BYTES);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_DEPTH);
    assert!(report.cell_count() <= MAX_CELLS);
    assert!(report.label_count() <= MAX_LABELS);
    assert!(report.text_bytes() <= MAX_TEXT_BYTES);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.retained_bytes(), source_bytes);
}

fn exercise_limit_failures(
    source: &[u8],
    route: Route,
    report: litchi_iwa_protos::chart_data_codec::DecodeReport,
) {
    let exact = DecodeOptions::new(
        report.source_bytes(),
        report.fields(),
        report.work_bytes(),
        report.max_depth(),
        report.cell_count(),
        report.label_count(),
        report.text_bytes(),
    );
    route
        .decode(source, &exact)
        .expect("observed chart-data limits must be inclusive");

    let check = |options: DecodeOptions, expected: fn(DecodeLimit) -> bool| {
        let error = route
            .decode(source, &options)
            .expect_err("one-below chart-data limit should reject the source");
        let limit = error
            .resource_limit()
            .expect("resource failure should identify a limit axis");
        assert!(expected(limit), "unexpected chart-data limit: {limit:?}");
        observe_error(error);
    };

    if report.source_bytes() > 0 {
        check(
            exact.with_max_input_bytes(report.source_bytes() - 1),
            |limit| matches!(limit, DecodeLimit::Bytes { .. }),
        );
    }
    if report.fields() > 0 {
        check(exact.with_max_fields(report.fields() - 1), |limit| {
            matches!(limit, DecodeLimit::Fields { .. })
        });
    }
    if report.work_bytes() > 0 {
        check(
            exact.with_max_work_bytes(report.work_bytes() - 1),
            |limit| matches!(limit, DecodeLimit::Work { .. }),
        );
    }
    if report.max_depth() > 0 {
        check(exact.with_max_depth(report.max_depth() - 1), |limit| {
            matches!(limit, DecodeLimit::Nesting { .. })
        });
    }
    if report.cell_count() > 0 {
        check(exact.with_max_cells(report.cell_count() - 1), |limit| {
            matches!(limit, DecodeLimit::Cells { .. })
        });
    }
    if report.label_count() > 0 {
        check(
            exact.with_max_label_count(report.label_count() - 1),
            |limit| matches!(limit, DecodeLimit::Labels { .. }),
        );
    }
    if report.text_bytes() > 0 {
        check(
            exact.with_max_text_bytes(report.text_bytes() - 1),
            |limit| matches!(limit, DecodeLimit::Text { .. }),
        );
    }
}

fn observe_error(error: DecodeError) {
    let report = error.report();
    black_box((
        error.resource_limit(),
        error.missing_required_field(),
        error.duplicate_singular_field(),
        error.input_limit_values(),
        error.field_limit_values(),
        error.work_limit_values(),
        error.cell_limit_values(),
        error.label_limit_values(),
        error.text_limit_values(),
        error.depth_limit_values(),
        error.is_non_finite_numeric(),
        error.is_non_rectangular(),
        report,
    ));
}

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

fn fixed64_field(field: u32, value: f64) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 1);
    output.extend_from_slice(&value.to_bits().to_le_bytes());
    output
}

fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 2);
    push_varint(
        &mut output,
        u64::try_from(payload.len()).expect("fixture payload length fits in u64"),
    );
    output.extend_from_slice(payload);
    output
}

fn joined(fields: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    fields.into_iter().flatten().collect()
}

fn grid_row(values: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    let payload = joined(values.into_iter().map(|value| length_field(1, &value)));
    length_field(3, &payload)
}

fn grid(
    row_names: &[&str],
    column_names: &[&str],
    rows: impl IntoIterator<Item = Vec<Vec<u8>>>,
    id_map: Option<&[u8]>,
) -> Vec<u8> {
    let mut output = Vec::new();
    for name in row_names {
        output.extend(length_field(1, name.as_bytes()));
    }
    for name in column_names {
        output.extend(length_field(2, name.as_bytes()));
    }
    for values in rows {
        output.extend(grid_row(values));
    }
    if let Some(id_map) = id_map {
        output.extend(length_field(4, id_map));
    }
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

fn fixed_cases() -> Vec<(Vec<u8>, Route)> {
    let valid_grid = grid(
        &["North", "South"],
        &["Q1", "Q2", "Q3"],
        [
            vec![
                joined([
                    varint_field(90, 17),
                    fixed64_field(1, 0.0),
                    fixed64_field(4, 45123.25),
                    length_field(89, b"cell future bytes"),
                ]),
                joined([
                    length_field(88, b"blank future bytes"),
                    fixed64_field(4, 45124.0),
                ]),
                fixed64_field(4, 45123.0),
            ],
            vec![
                joined([
                    fixed64_field(3, 0.5),
                    fixed64_field(1, -0.0),
                    varint_field(87, 29),
                ]),
                joined([fixed64_field(3, 3.5), fixed64_field(1, -9.25)]),
                fixed64_field(1, 42.0),
            ],
        ],
        Some(&[0xff, 0x80, 0x80, 0x01]),
    );
    let valid_modern = modern(&valid_grid);

    let ragged = grid(
        &["North", "South"],
        &["Q1", "Q2"],
        [
            vec![fixed64_field(1, 1.0), fixed64_field(1, 2.0)],
            vec![fixed64_field(1, 3.0)],
        ],
        None,
    );
    let non_finite = grid(
        &["North"],
        &["Q1"],
        [vec![fixed64_field(1, f64::NAN)]],
        None,
    );
    let duplicate = grid(
        &["North"],
        &["Q1"],
        [vec![joined([fixed64_field(1, 0.0), fixed64_field(1, 4.0)])]],
        None,
    );
    let malformed = joined([
        // Truncated length-delimited grid row.
        vec![0x1a, 0x04, 0x0a],
        // Wrong wire type for a row name.
        varint_field(1, 1),
        // Unterminated length field.
        vec![0x12, 0x80],
    ]);

    vec![
        (valid_grid, Route::Grid),
        (valid_modern, Route::Modern),
        (ragged, Route::Grid),
        (non_finite, Route::Grid),
        (duplicate, Route::Grid),
        (malformed, Route::Grid),
    ]
}
