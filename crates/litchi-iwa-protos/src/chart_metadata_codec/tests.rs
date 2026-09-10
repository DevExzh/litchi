use std::ptr;

use super::{
    ChartMetadataFormat, DecodeError, DecodeLimit, DecodeOptions, decode_legacy,
    decode_legacy_with_report, decode_modern, decode_modern_with_report,
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

fn signed_field(field: u32, value: i32) -> Vec<u8> {
    varint_field(field, i64::from(value) as u64)
}

fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 2);
    push_varint(&mut output, payload.len() as u64);
    output.extend_from_slice(payload);
    output
}

fn bool_field(field: u32, value: bool) -> Vec<u8> {
    varint_field(field, u64::from(value))
}

fn append(output: &mut Vec<u8>, fields: impl IntoIterator<Item = Vec<u8>>) {
    for field in fields {
        output.extend(field);
    }
}

fn joined(fields: impl IntoIterator<Item = Vec<u8>>) -> Vec<u8> {
    let mut output = Vec::new();
    append(&mut output, fields);
    output
}

fn reference(identifier: u64) -> Vec<u8> {
    varint_field(1, identifier)
}

fn group(field: u32, contents: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, u64::from(field) << 3 | 3);
    output.extend_from_slice(contents);
    push_varint(&mut output, u64::from(field) << 3 | 4);
    output
}

fn modern_source() -> (Vec<u8>, Vec<u8>) {
    let mut grid = Vec::new();
    append(
        &mut grid,
        [
            length_field(1, b"Q1"),
            length_field(1, b"Q2"),
            length_field(2, b"North"),
            length_field(3, &varint_field(1, 1)),
            length_field(3, &varint_field(1, 2)),
        ],
    );

    let mut chart = Vec::new();
    append(
        &mut chart,
        [
            signed_field(1, -1_234),
            bool_field(6, false),
            length_field(7, &grid),
            length_field(10, &reference(77)),
        ],
    );

    let mut source = length_field(10_000, &chart);
    source.extend(varint_field(99, 0xfeed));
    (source, grid)
}

fn modern_without_grid() -> Vec<u8> {
    length_field(10_000, &signed_field(1, 28))
}

fn legacy_source_with_grid() -> (Vec<u8>, Vec<u8>) {
    let mut grid = Vec::new();
    append(
        &mut grid,
        [
            signed_field(1, -17),
            length_field(2, "行".as_bytes()),
            length_field(3, "North".as_bytes()),
            length_field(4, &varint_field(1, 1)),
            length_field(4, &varint_field(1, 2)),
        ],
    );

    let mut model = Vec::new();
    append(
        &mut model,
        [length_field(2, &reference(41)), length_field(5, &grid)],
    );

    let mut source = Vec::new();
    append(
        &mut source,
        [
            length_field(2, &model),
            signed_field(4, -5_678),
            length_field(14, &reference(88)),
        ],
    );
    source.extend(varint_field(99, 0xbeef));
    (source, grid)
}

fn legacy_without_inline_grid() -> Vec<u8> {
    let mut model = Vec::new();
    // These are deliberately opaque and malformed as TSP.Reference payloads.
    // The legacy model.grid edge is not part of the metadata projection.
    append(
        &mut model,
        [length_field(2, &[0xff]), length_field(2, &[0x80, 0x80])],
    );
    let mut source = length_field(2, &model);
    source.extend(signed_field(4, i32::MIN));
    source
}

fn permissive_options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(source.len(), 1_000_000, 1_000_000, 64, 1_000_000, 1_000_000)
}

fn embedded_subslice<'source>(source: &'source [u8], needle: &[u8]) -> &'source [u8] {
    let offset = source
        .windows(needle.len())
        .position(|window| window == needle)
        .expect("fixture contains embedded payload");
    &source[offset..offset + needle.len()]
}

fn assert_resource_limit(error: &DecodeError, expected: DecodeLimit, label: &str) {
    assert_eq!(error.resource_limit(), Some(expected), "{label}: {error}");
    assert_eq!(error.report().failure_work_bytes(), 0, "{label}: {error}");
}

#[test]
fn modern_labels_kind_default_and_nested_source_are_borrowed() {
    let (source, grid) = modern_source();
    let options = permissive_options(&source);
    let (snapshot, report) = decode_modern_with_report(&source, &options).expect("modern chart");

    assert_eq!(snapshot.format(), ChartMetadataFormat::Modern);
    assert_eq!(snapshot.chart_type(), -1_234);
    assert_eq!(snapshot.contains_default_data(), Some(false));
    assert_eq!(
        snapshot.row_labels().iter().collect::<Vec<_>>(),
        ["Q1", "Q2"]
    );
    assert_eq!(
        snapshot.column_labels().iter().collect::<Vec<_>>(),
        ["North"]
    );
    assert_eq!(snapshot.series_count(), 2);
    assert_eq!(snapshot.non_style_ref().map(|id| id.get()), Some(77));
    assert!(ptr::eq(snapshot.source(), source.as_slice()));

    let embedded_grid = embedded_subslice(&source, &grid);
    assert!(ptr::eq(
        snapshot.grid_source().expect("modern grid"),
        embedded_grid
    ));
    let row = snapshot.row_labels().get(0).expect("first row label");
    let expected_row = std::str::from_utf8(embedded_subslice(&source, b"Q1")).unwrap();
    assert!(ptr::eq(row, expected_row));
    assert_eq!(report.source_bytes(), source.len());
    assert_eq!(report.retained_bytes(), source.len());
    assert_eq!(report.allocations(), 0);
}

#[test]
fn modern_absent_optional_fields_are_not_synthesized() {
    let source = modern_without_grid();
    let snapshot = decode_modern(&source, &permissive_options(&source)).expect("modern chart");

    assert_eq!(snapshot.format(), ChartMetadataFormat::Modern);
    assert_eq!(snapshot.chart_type(), 28);
    assert_eq!(snapshot.contains_default_data(), None);
    assert_eq!(snapshot.row_labels().len(), 0);
    assert_eq!(snapshot.column_labels().len(), 0);
    assert_eq!(snapshot.series_count(), 0);
    assert_eq!(snapshot.non_style_ref(), None);
    assert_eq!(snapshot.grid_source(), None);
}

#[test]
fn legacy_labels_direction_kind_and_grid_source_are_borrowed() {
    let (source, grid) = legacy_source_with_grid();
    let snapshot = decode_legacy(&source, &permissive_options(&source)).expect("legacy chart");

    assert_eq!(snapshot.format(), ChartMetadataFormat::Legacy);
    assert_eq!(snapshot.chart_type(), -5_678);
    assert_eq!(snapshot.contains_default_data(), None);
    assert_eq!(snapshot.row_labels().iter().collect::<Vec<_>>(), ["行"]);
    assert_eq!(
        snapshot.column_labels().iter().collect::<Vec<_>>(),
        ["North"]
    );
    assert_eq!(snapshot.series_count(), 2);
    assert_eq!(snapshot.non_style_ref().map(|id| id.get()), Some(88));
    assert!(ptr::eq(snapshot.source(), source.as_slice()));
    assert!(ptr::eq(
        snapshot.grid_source().expect("legacy inline grid"),
        embedded_subslice(&source, &grid)
    ));
}

#[test]
fn legacy_missing_inline_grid_has_zero_series_and_opaque_duplicate_edges() {
    let source = legacy_without_inline_grid();
    let (snapshot, report) =
        decode_legacy_with_report(&source, &permissive_options(&source)).expect("legacy chart");

    assert_eq!(snapshot.chart_type(), i32::MIN);
    assert_eq!(snapshot.grid_source(), None);
    assert_eq!(snapshot.series_count(), 0);
    assert!(snapshot.row_labels().is_empty());
    assert!(snapshot.column_labels().is_empty());
    assert_eq!(snapshot.non_style_ref(), None);
    assert_eq!(report.allocations(), 0);
    assert_eq!(report.source_bytes(), source.len());
}

#[test]
fn zero_reference_is_absent_but_duplicate_nonzero_reference_is_rejected() {
    let mut chart = Vec::new();
    append(
        &mut chart,
        [signed_field(1, 7), length_field(10, &reference(0))],
    );
    let zero_source = length_field(10_000, &chart);
    let snapshot = decode_modern(&zero_source, &permissive_options(&zero_source))
        .expect("zero reference is a valid absent edge");
    assert_eq!(snapshot.non_style_ref(), None);

    let duplicate = [length_field(
        10_000,
        &joined([
            signed_field(1, 7),
            length_field(10, &reference(1)),
            length_field(10, &reference(2)),
        ]),
    )]
    .concat();
    let error = decode_modern(&duplicate, &permissive_options(&duplicate)).unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartArchive.chart_non_style")
    );

    for second_identifier in [0, 2] {
        let reference_payload = joined([varint_field(1, 0), varint_field(1, second_identifier)]);
        let source = length_field(
            10_000,
            &joined([signed_field(1, 7), length_field(10, &reference_payload)]),
        );
        let error = decode_modern(&source, &permissive_options(&source)).unwrap_err();
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier"),
            "second identifier {second_identifier} must not hide the duplicate",
        );

        let legacy = joined([
            length_field(2, &[]),
            signed_field(4, 7),
            length_field(14, &reference_payload),
        ]);
        let error = decode_legacy(&legacy, &permissive_options(&legacy)).unwrap_err();
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier"),
            "legacy second identifier {second_identifier} must not hide the duplicate",
        );
    }
}

#[test]
fn duplicate_selected_fields_are_rejected_without_a_snapshot() {
    let duplicate_chart_type = [length_field(
        10_000,
        &joined([signed_field(1, 1), signed_field(1, 2)]),
    )]
    .concat();
    let error = decode_modern(
        &duplicate_chart_type,
        &permissive_options(&duplicate_chart_type),
    )
    .unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartArchive.chart_type")
    );

    let duplicate_default = [length_field(
        10_000,
        &joined([bool_field(6, false), bool_field(6, true)]),
    )]
    .concat();
    let error =
        decode_modern(&duplicate_default, &permissive_options(&duplicate_default)).unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartArchive.contains_default_data")
    );

    let duplicate_grid = [length_field(
        10_000,
        &joined([length_field(7, &[]), length_field(7, &[])]),
    )]
    .concat();
    let error = decode_modern(&duplicate_grid, &permissive_options(&duplicate_grid)).unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartArchive.grid")
    );

    let duplicate_extension = [
        length_field(10_000, &signed_field(1, 1)),
        length_field(10_000, &signed_field(1, 2)),
    ]
    .concat();
    let error = decode_modern(
        &duplicate_extension,
        &permissive_options(&duplicate_extension),
    )
    .unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.ChartDrawableArchive.unity")
    );

    let mut model = Vec::new();
    append(&mut model, [length_field(5, &[]), length_field(5, &[])]);
    let legacy = [length_field(2, &model), signed_field(4, 1)].concat();
    let error = decode_legacy(&legacy, &permissive_options(&legacy)).unwrap_err();
    assert_eq!(
        error.duplicate_singular_field(),
        Some("TSCH.PreUFF.ChartModelArchive.inline_grid")
    );
}

#[test]
fn malformed_utf8_wrong_wire_and_missing_reference_fail_closed() {
    let invalid_grid = length_field(1, &[0xff]);
    let invalid_chart = length_field(
        10_000,
        &joined([signed_field(1, 1), length_field(7, &invalid_grid)]),
    );
    let error = decode_modern(&invalid_chart, &permissive_options(&invalid_chart)).unwrap_err();
    assert_eq!(
        error.invalid_utf8_field(),
        Some("TSCH.ChartGridArchive.row_name")
    );

    let wrong_wire_chart = length_field(10_000, &joined([length_field(1, b"wrong wire")]));
    assert!(decode_modern(&wrong_wire_chart, &permissive_options(&wrong_wire_chart)).is_err());

    let missing_reference =
        length_field(10_000, &joined([signed_field(1, 1), length_field(10, &[])]));
    let error =
        decode_modern(&missing_reference, &permissive_options(&missing_reference)).unwrap_err();
    assert_eq!(
        error.missing_required_field(),
        Some("TSP.Reference.identifier")
    );
}

#[test]
fn noncanonical_values_and_malformed_groups_are_rejected() {
    // TSCH.ChartArchive.chart_type = 1 with a non-minimal varint value.
    let noncanonical_value = length_field(10_000, &[0x08, 0x81, 0x00]);
    let error = decode_modern(
        &noncanonical_value,
        &permissive_options(&noncanonical_value),
    )
    .unwrap_err();
    assert_eq!(error.noncanonical_reason(), Some("protobuf varint value"));

    let noncanonical_key_chart = vec![0x88, 0x00, 0x01];
    let source = length_field(10_000, &noncanonical_key_chart);
    let error = decode_modern(&source, &permissive_options(&source)).unwrap_err();
    assert_eq!(error.noncanonical_reason(), Some("protobuf field key"));

    let valid_group = [
        length_field(10_000, &signed_field(1, 1)),
        group(99, &varint_field(100, 7)),
    ]
    .concat();
    assert!(decode_modern(&valid_group, &permissive_options(&valid_group)).is_ok());

    let balanced_group = group(99, &varint_field(100, 7));
    let malformed_group = [
        length_field(10_000, &signed_field(1, 1)),
        balanced_group[..balanced_group.len() - 1].to_vec(),
        varint_field(100, 7),
    ]
    .concat();
    assert!(decode_modern(&malformed_group, &permissive_options(&malformed_group)).is_err());
}

#[test]
fn exact_and_one_below_limits_report_the_attempted_resource() {
    let (source, _) = modern_source();
    let broad = permissive_options(&source);
    let (_, report) = decode_modern_with_report(&source, &broad).expect("baseline chart");

    let exact_input = broad.with_max_input_bytes(report.source_bytes());
    assert!(decode_modern(&source, &exact_input).is_ok());
    let error = decode_modern(
        &source,
        &exact_input.with_max_input_bytes(report.source_bytes() - 1),
    )
    .unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Bytes {
            observed: report.source_bytes(),
            maximum: report.source_bytes() - 1,
        },
        "input bytes",
    );

    let exact_fields = broad.with_max_fields(report.fields());
    assert!(decode_modern(&source, &exact_fields).is_ok());
    let error =
        decode_modern(&source, &exact_fields.with_max_fields(report.fields() - 1)).unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Fields {
            observed: report.fields(),
            maximum: report.fields() - 1,
        },
        "fields",
    );

    let exact_work = broad.with_max_work_bytes(report.work_bytes());
    assert!(decode_modern(&source, &exact_work).is_ok());
    let error = decode_modern(
        &source,
        &exact_work.with_max_work_bytes(report.work_bytes() - 1),
    )
    .unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Work {
            observed: report.work_bytes(),
            maximum: report.work_bytes() - 1,
        },
        "work",
    );

    let exact_labels = broad.with_max_label_count(report.label_count());
    assert!(decode_modern(&source, &exact_labels).is_ok());
    let error = decode_modern(
        &source,
        &exact_labels.with_max_label_count(report.label_count() - 1),
    )
    .unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Labels {
            observed: report.label_count(),
            maximum: report.label_count() - 1,
        },
        "labels",
    );

    let exact_text = broad.with_max_text_bytes(report.text_bytes());
    assert!(decode_modern(&source, &exact_text).is_ok());
    let error = decode_modern(
        &source,
        &exact_text.with_max_text_bytes(report.text_bytes() - 1),
    )
    .unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Text {
            observed: report.text_bytes(),
            maximum: report.text_bytes() - 1,
        },
        "text",
    );

    let exact_depth = broad.with_max_depth(report.max_depth());
    assert!(decode_modern(&source, &exact_depth).is_ok());
    let error =
        decode_modern(&source, &exact_depth.with_max_depth(report.max_depth() - 1)).unwrap_err();
    assert_resource_limit(
        &error,
        DecodeLimit::Nesting {
            observed: report.max_depth(),
            maximum: report.max_depth() - 1,
        },
        "nesting",
    );
}
