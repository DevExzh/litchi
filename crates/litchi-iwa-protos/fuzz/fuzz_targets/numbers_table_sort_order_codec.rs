#![no_main]

//! Strict bounded fuzzing for `TST.TableModelArchive.sort_order` (field 44).
//!
//! The target keeps the complete model payload source-authoritative, compares
//! scalar and reported decode routes, exercises malformed/unknown-group
//! inputs, and replays every prepared rewrite requirement with exact and
//! max-minus-one execution ceilings.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::table_sort_order_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_RULES: usize = 1_024;
const MAX_COLUMNS: usize = 4 * 1_024;

const UNKNOWN_GROUP: &[u8] = &[0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06];

const FIXED_CASES: &[&[u8]] = &[
    // Empty TableModelArchive: field 44 is absent and decodes as None.
    &[],
    // field 44 = { type: EntireTable, rule(index=0, ascending) }.
    &[
        0xa2, 0x02, 0x08, 0x08, 0x00, 0x12, 0x04, 0x08, 0x00, 0x10, 0x00,
    ],
    // field 44 = { type: SelectedRows, rules(index=1 descending, 2 ascending) }.
    &[
        0xa2, 0x02, 0x0e, 0x08, 0x01, 0x12, 0x04, 0x08, 0x01, 0x10, 0x01, 0x12, 0x04, 0x08, 0x02,
        0x10, 0x00,
    ],
    // Unknown overlong scalar and balanced unknown group retained around a
    // valid sort rule.
    &[
        0xa2, 0x02, 0x12, 0x08, 0x00, 0x12, 0x04, 0x08, 0x00, 0x10, 0x00, 0x98, 0x06, 0x81, 0x00,
        0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06,
    ],
    // Duplicate required sort scope.
    &[0xa2, 0x02, 0x04, 0x08, 0x00, 0x08, 0x00],
    // Wrong wire type for the nested sort scope.
    &[0xa2, 0x02, 0x02, 0x0d, 0x00],
    // Non-canonical known scope varint.
    &[0xa2, 0x02, 0x03, 0x08, 0x80, 0x00],
    // Truncated nested rule.
    &[0xa2, 0x02, 0x06, 0x08, 0x00, 0x12, 0x04, 0x08, 0x00],
    // Duplicate rule column.
    &[
        0xa2, 0x02, 0x0c, 0x08, 0x00, 0x12, 0x04, 0x08, 0x00, 0x10, 0x00, 0x12, 0x04, 0x08, 0x00,
        0x10, 0x01,
    ],
    // Unterminated unknown group.
    &[
        0xa2, 0x02, 0x0e, 0x08, 0x00, 0x12, 0x04, 0x08, 0x00, 0x10, 0x00, 0x9b, 0x06, 0x08, 0x01,
    ],
];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };
    exercise_source(&source, data);

    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        for fixed in FIXED_CASES {
            exercise_source(fixed, b"fixed-table-sort-order");
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

fn options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        MAX_INPUT_BYTES.max(source.len()),
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_RULES,
        MAX_COLUMNS,
    )
    .with_max_allocations(16)
}

fn exercise_source(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let decode_options = options(source);
    let scalar = codec::decode_table_sort_order(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "sort scalar decode modified source"
    );
    let reported = codec::decode_table_sort_order_with_report(source, decode_options);
    assert_eq!(
        source,
        before.as_slice(),
        "sort report decode modified source"
    );

    match (scalar, reported) {
        (Ok(snapshot), Ok((reported_snapshot, report))) => {
            assert_eq!(snapshot, reported_snapshot);
            assert!(report.input_bytes() <= MAX_INPUT_BYTES.max(source.len()));
            assert!(report.output_bytes() <= MAX_OUTPUT_BYTES.max(source.len()));
            assert!(report.fields() <= MAX_FIELDS);
            assert!(report.work_bytes() <= MAX_WORK_BYTES);
            assert!(report.max_depth() <= MAX_RECURSION);
            assert!(report.rules() <= MAX_RULES);
            assert!(report.allocations() <= 16);
            black_box((&snapshot, report));
            exercise_rewrites(source, snapshot, data, &before);
        },
        (Err(scalar_error), Err(reported_error)) => {
            observe_error(scalar_error);
            observe_error(reported_error);
        },
        (scalar_result, reported_result) => panic!(
            "sort scalar/report decode disagreement: scalar={:?}, report={:?}",
            scalar_result
                .as_ref()
                .map(|value| value.as_ref().map(|order| order.rules().len())),
            reported_result
                .as_ref()
                .map(|(order, _)| order.as_ref().map(|order| order.rules().len()))
        ),
    }

    assert_eq!(source, before.as_slice(), "sort probes modified source");
}

fn exercise_rewrites(
    source: &[u8],
    current: Option<codec::SortOrderSnapshot>,
    data: &[u8],
    before: &[u8],
) {
    let desired = desired_order(data);
    let requests = [current.clone(), desired.clone(), None];
    for requested in requests {
        let prepared =
            codec::prepare_table_sort_order_rewrite(source, requested.clone(), options(source));
        let Ok(plan) = prepared else {
            continue;
        };
        let requirements = plan.execution_requirements();
        let preparation = plan.prepare_report();
        assert_eq!(preparation.output_bytes(), requirements.output_bytes);
        assert_eq!(preparation.fields(), requirements.fields);
        assert_eq!(preparation.work_bytes(), requirements.work_bytes);
        assert_eq!(preparation.max_depth(), requirements.max_depth);
        assert_eq!(preparation.rules(), requirements.rules);
        assert_eq!(preparation.allocations(), requirements.allocations);
        assert_eq!(preparation.retained_bytes(), requirements.retained_bytes);
        assert_eq!(preparation.scratch_bytes(), requirements.scratch_bytes);

        let exact = plan
            .execute(codec::RewriteExecutionLimits::exact(requirements))
            .unwrap_or_else(|error| panic!("exact sort prepared replay failed: {error}"));
        assert_eq!(source, before, "sort rewrite modified source");
        assert_eq!(exact.report().output_bytes(), exact.bytes().len());
        assert_eq!(exact.report().fields(), requirements.fields);
        assert_eq!(exact.report().work_bytes(), requirements.work_bytes);
        assert_eq!(exact.report().rules(), requirements.rules);
        assert_eq!(
            codec::decode_table_sort_order(exact.bytes(), options(exact.bytes()))
                .unwrap_or_else(|error| panic!("sort candidate readback failed: {error}")),
            requested
        );

        let one_shot = codec::rewrite_table_sort_order(source, requested.clone(), options(source))
            .unwrap_or_else(|error| panic!("sort one-shot rewrite failed: {error}"));
        assert_eq!(one_shot.bytes(), exact.bytes());
        assert_eq!(one_shot.report(), exact.report());

        exercise_limit_failures(source, requested.clone(), requirements);
        if source
            .windows(UNKNOWN_GROUP.len())
            .any(|window| window == UNKNOWN_GROUP)
            && requested.is_some()
            && requested != current
        {
            assert!(
                exact
                    .bytes()
                    .windows(UNKNOWN_GROUP.len())
                    .any(|window| window == UNKNOWN_GROUP)
            );
        }
        black_box(exact);
    }
}

fn exercise_limit_failures(
    source: &[u8],
    desired: Option<codec::SortOrderSnapshot>,
    requirements: codec::RewriteExecutionRequirements,
) {
    let probes = [
        (requirements.output_bytes > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_output_bytes(requirements.output_bytes.saturating_sub(1))
        }),
        (requirements.fields > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_fields(requirements.fields.saturating_sub(1))
        }),
        (requirements.work_bytes > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_work_bytes(requirements.work_bytes.saturating_sub(1))
        }),
        (requirements.max_depth > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_max_depth(requirements.max_depth.saturating_sub(1))
        }),
        (requirements.rules > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_rules(requirements.rules.saturating_sub(1))
        }),
        (requirements.allocations > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_allocations(requirements.allocations.saturating_sub(1))
        }),
        (requirements.retained_bytes > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_retained_bytes(requirements.retained_bytes.saturating_sub(1))
        }),
        (requirements.scratch_bytes > 0).then(|| {
            codec::RewriteExecutionLimits::exact(requirements)
                .with_scratch_bytes(requirements.scratch_bytes.saturating_sub(1))
        }),
    ];
    for limits in probes.into_iter().flatten() {
        let plan =
            codec::prepare_table_sort_order_rewrite(source, desired.clone(), options(source));
        let Ok(plan) = plan else {
            continue;
        };
        let result = plan.execute(limits);
        assert!(
            result.is_err(),
            "sort rewrite accepted a max-minus-one limit"
        );
        if let Err(error) = result {
            black_box(error.resource_limit());
        }
    }
}

fn desired_order(data: &[u8]) -> Option<codec::SortOrderSnapshot> {
    let scope = if data.get(1).copied().unwrap_or_default() & 1 == 0 {
        codec::SortScope::EntireTable
    } else {
        codec::SortScope::SelectedRows
    };
    let count = usize::from(data.first().copied().unwrap_or_default() % 3) + 1;
    let rules = (0..count)
        .map(|index| {
            let column = u32::from(data.get(index + 2).copied().unwrap_or(index as u8))
                % u32::try_from(MAX_COLUMNS).unwrap_or(u32::MAX);
            let direction = if data.get(index + count + 2).copied().unwrap_or_default() & 1 == 0 {
                codec::SortDirection::Ascending
            } else {
                codec::SortDirection::Descending
            };
            codec::SortRule::new(column, direction)
        })
        .collect::<Vec<_>>();
    codec::SortOrderSnapshot::new(scope, rules).ok()
}

fn observe_error(error: impl std::fmt::Debug + std::fmt::Display) {
    black_box(error.to_string());
    black_box(format_args!("{error:?}"));
}
