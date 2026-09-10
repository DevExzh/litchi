#![no_main]

//! Bounded raw and structured fuzzing for the native Numbers table-merge
//! reader.
//!
//! Arbitrary input is admitted only through a one-kilobyte slice.  The same
//! input also selects small hand-written TableModelArchive wire fixtures so
//! the merge-owner path remains reachable even when random bytes are refused
//! before the selected model.  The fixtures use the canonical native formula
//! shape directly; no generated AST values cross this fuzz boundary.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_common::{Error, WireLimits, table::merge::Region};
use litchi_numbers_wire::table_merges::{self, ReadLimits};

const MAX_INPUT_BYTES: usize = 1024;
const MAX_FIELDS: usize = 4096;
const MAX_REWRITE_WORK: usize = 16 * 1024;
const MAX_REGIONS: usize = 16;
const MAX_OVERLAP_CHECKS: usize = 128;
const TABLE_ROWS: u32 = 8;
const TABLE_COLUMNS: u32 = 8;

// Parsing this value yields owner CFUUID words [1, 0, 2, 0] after the native
// table-UUID byte swap performed by the shared wire adapter.
const TABLE_ID: &[u8] = b"01000000000000000200000000000000";

#[derive(Clone, Copy, Debug)]
struct Rectangle {
    begin_row: u32,
    begin_column: u32,
    end_row: u32,
    end_column: u32,
}

#[derive(Clone, Copy, Debug)]
enum Fixture {
    Valid,
    InvalidProfile,
    InvalidIndex,
    ForeignUuid,
    Overlap,
    Bounds,
    NonCanonicalKey,
    NonCanonicalLength,
}

const FIXTURES: [Fixture; 8] = [
    Fixture::Valid,
    Fixture::InvalidProfile,
    Fixture::InvalidIndex,
    Fixture::ForeignUuid,
    Fixture::Overlap,
    Fixture::Bounds,
    Fixture::NonCanonicalKey,
    Fixture::NonCanonicalLength,
];

fuzz_target!(|data: &[u8]| {
    let raw = &data[..data.len().min(MAX_INPUT_BYTES)];
    exercise_raw(raw);

    // Run every small structured branch for each input.  This keeps the
    // campaign from depending on a particular command-byte distribution to
    // reach the canonical-field, identity, overlap, and bounds checks.
    for fixture in FIXTURES {
        exercise_fixture(fixture, data);
    }
});

fn exercise_raw(source: &[u8]) {
    let limits = fuzz_limits();
    match table_merges::read_table_merges(source, limits) {
        Ok(read) => {
            assert_report_within(read.report, limits);
            black_box((read.regions.len(), read.report));
        },
        Err(error) => {
            assert_attempted_within(error.attempted(), limits);
            black_box((error.error(), error.attempted()));
        },
    }
}

fn exercise_fixture(fixture: Fixture, data: &[u8]) {
    let source = fixture_source(fixture, data);
    let limits = fuzz_limits();

    match fixture {
        Fixture::Valid => {
            let read = table_merges::read_table_merges(&source, limits)
                .unwrap_or_else(|error| panic!("canonical merge fixture was rejected: {error}"));
            let expected = Region::new(0, 0, 2, 2)
                .unwrap_or_else(|error| panic!("canonical fixture region is invalid: {error}"));
            assert_eq!(read.regions, [expected]);
            assert_report_within(read.report, limits);
            black_box(read);
        },
        Fixture::InvalidProfile => exercise_invalid_profiles(&source),
        Fixture::InvalidIndex => expect_invalid(&source, limits, "formula index"),
        Fixture::ForeignUuid => expect_invalid(&source, limits, "foreign table UUID"),
        Fixture::Overlap => expect_invalid(&source, limits, "overlapping regions"),
        Fixture::Bounds => expect_invalid(&source, limits, "out-of-bounds region"),
        Fixture::NonCanonicalKey => expect_invalid(&source, limits, "non-canonical field key"),
        Fixture::NonCanonicalLength => {
            expect_invalid(&source, limits, "non-canonical field length")
        },
    }
}

fn exercise_invalid_profiles(source: &[u8]) {
    let mut too_many_regions = fuzz_limits();
    too_many_regions.max_regions = WireLimits::MAX_FIELDS + 1;
    let error = table_merges::read_table_merges(source, too_many_regions)
        .expect_err("an over-ceiling region profile must be rejected");
    assert!(matches!(error.error(), Error::InvalidLimit { .. }));
    assert_eq!(error.attempted().input_bytes, 0);
    assert_eq!(error.attempted().fields, 0);
    assert_eq!(error.attempted().work, 0);
    black_box(error);

    let mut too_much_overlap_work = fuzz_limits();
    too_much_overlap_work.max_overlap_checks = WireLimits::MAX_REWRITE_WORK + 1;
    let error = table_merges::read_table_merges(source, too_much_overlap_work)
        .expect_err("an over-ceiling overlap profile must be rejected");
    assert!(matches!(error.error(), Error::InvalidLimit { .. }));
    assert_eq!(error.attempted().input_bytes, 0);
    assert_eq!(error.attempted().fields, 0);
    assert_eq!(error.attempted().work, 0);
    black_box(error);
}

fn expect_invalid(source: &[u8], limits: ReadLimits, label: &str) {
    let error = match table_merges::read_table_merges(source, limits) {
        Ok(_) => panic!("{label} fixture unexpectedly succeeded"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::InvalidFormat(_) | Error::LimitExceeded { .. } | Error::Allocation { .. }
    ));
    assert_attempted_within(error.attempted(), limits);
    black_box(error);
}

fn fuzz_limits() -> ReadLimits {
    let wire = WireLimits::default()
        .with_input_bytes(MAX_INPUT_BYTES)
        .unwrap_or_else(|error| panic!("valid fuzz input profile: {error}"))
        .with_fields(MAX_FIELDS)
        .unwrap_or_else(|error| panic!("valid fuzz field profile: {error}"))
        .with_nesting(16)
        .unwrap_or_else(|error| panic!("valid fuzz nesting profile: {error}"))
        .with_rewrite_work(MAX_REWRITE_WORK)
        .unwrap_or_else(|error| panic!("valid fuzz work profile: {error}"));
    ReadLimits {
        wire,
        max_regions: MAX_REGIONS,
        max_overlap_checks: MAX_OVERLAP_CHECKS,
    }
}

fn assert_report_within(report: table_merges::ReadReport, limits: ReadLimits) {
    assert!(report.input_bytes() <= limits.wire.max_input_bytes());
    assert!(report.fields() <= limits.wire.max_fields());
    assert!(report.work() <= limits.wire.max_rewrite_work());
}

fn assert_attempted_within(cost: table_merges::AttemptedCost, limits: ReadLimits) {
    assert!(cost.input_bytes <= limits.wire.max_input_bytes());
    assert!(cost.fields <= limits.wire.max_fields());
    assert!(cost.work <= limits.wire.max_rewrite_work());
}

fn fixture_source(fixture: Fixture, data: &[u8]) -> Vec<u8> {
    let (rows, columns, rectangles, indices, next_index, foreign) = match fixture {
        Fixture::Valid
        | Fixture::InvalidProfile
        | Fixture::ForeignUuid
        | Fixture::NonCanonicalKey
        | Fixture::NonCanonicalLength => (
            TABLE_ROWS,
            TABLE_COLUMNS,
            vec![Rectangle {
                begin_row: 0,
                begin_column: 0,
                end_row: 1,
                end_column: 1,
            }],
            vec![1],
            2u32,
            matches!(fixture, Fixture::ForeignUuid),
        ),
        Fixture::InvalidIndex => (
            TABLE_ROWS,
            TABLE_COLUMNS,
            vec![Rectangle {
                begin_row: 0,
                begin_column: 0,
                end_row: 1,
                end_column: 1,
            }],
            vec![2],
            2u32,
            false,
        ),
        Fixture::Overlap => (
            TABLE_ROWS,
            TABLE_COLUMNS,
            vec![
                Rectangle {
                    begin_row: 0,
                    begin_column: 0,
                    end_row: 1,
                    end_column: 1,
                },
                Rectangle {
                    begin_row: 1,
                    begin_column: 1,
                    end_row: 2,
                    end_column: 2,
                },
            ],
            vec![1, 2],
            3u32,
            false,
        ),
        Fixture::Bounds => (
            2u32,
            2,
            vec![Rectangle {
                begin_row: 0,
                begin_column: 0,
                end_row: 2,
                end_column: 1,
            }],
            vec![1],
            2,
            false,
        ),
    };

    let mut model = Vec::new();
    match fixture {
        Fixture::NonCanonicalKey => bytes_field_with_framing(&mut model, 1, TABLE_ID, false, true),
        Fixture::NonCanonicalLength => {
            bytes_field_with_framing(&mut model, 1, TABLE_ID, true, false)
        },
        _ => bytes_field(&mut model, 1, TABLE_ID),
    }
    varint_field(&mut model, 6, u64::from(rows));
    varint_field(&mut model, 7, u64::from(columns));

    let mut store = Vec::new();
    varint_field(&mut store, 2, u64::from(next_index));
    let uid = if foreign { [1, 0, 3, 0] } else { [1, 0, 2, 0] };
    for (rectangle, index) in rectangles.iter().copied().zip(indices.iter().copied()) {
        let pair = formula_pair(index, rectangle, uid);
        bytes_field(&mut store, 3, &pair);
    }
    let mut owner = Vec::new();
    bytes_field(&mut owner, 2, &store);
    bytes_field(&mut model, 47, &owner);

    // One byte selects whether the source is copied with an opaque unknown
    // root field.  This keeps unknown-field admission exercised without ever
    // growing the structured fixture beyond the bounded profile.
    if data.first().copied().unwrap_or_default() & 1 == 1 {
        varint_field(&mut model, 99, 1);
    }
    model
}

fn formula_pair(index: u32, rectangle: Rectangle, uid: [u32; 4]) -> Vec<u8> {
    let mut pair = Vec::new();
    varint_field(&mut pair, 1, u64::from(index));
    let formula = merge_formula(rectangle, uid);
    bytes_field(&mut pair, 2, &formula);
    pair
}

fn merge_formula(rectangle: Rectangle, uid: [u32; 4]) -> Vec<u8> {
    let mut cfuuid = Vec::new();
    for (field, word) in (2..=5).zip(uid) {
        varint_field(&mut cfuuid, field, u64::from(word));
    }
    let mut cross_extra = Vec::new();
    bytes_field(&mut cross_extra, 1, &cfuuid);

    let mut sticky = Vec::new();
    for field in 1..=4 {
        varint_field(&mut sticky, field, 1);
    }

    let mut tract = Vec::new();
    let column_range = absolute_range(rectangle.begin_column, rectangle.end_column);
    let row_range = absolute_range(rectangle.begin_row, rectangle.end_row);
    bytes_field(&mut tract, 3, &column_range);
    bytes_field(&mut tract, 4, &row_range);
    varint_field(&mut tract, 5, 1);

    let mut range_node = Vec::new();
    varint_field(&mut range_node, 1, 67);
    bytes_field(&mut range_node, 28, &cross_extra);
    bytes_field(&mut range_node, 33, &sticky);
    bytes_field(&mut range_node, 40, &tract);

    let mut function_node = Vec::new();
    varint_field(&mut function_node, 1, 16);
    varint_field(&mut function_node, 2, 168);
    varint_field(&mut function_node, 3, 1);

    let mut ast = Vec::new();
    bytes_field(&mut ast, 1, &range_node);
    bytes_field(&mut ast, 1, &function_node);
    let mut formula = Vec::new();
    bytes_field(&mut formula, 1, &ast);
    formula
}

fn absolute_range(begin: u32, end: u32) -> Vec<u8> {
    let mut range = Vec::new();
    varint_field(&mut range, 1, u64::from(begin));
    if end != begin {
        varint_field(&mut range, 2, u64::from(end));
    }
    range
}

fn bytes_field(output: &mut Vec<u8>, field: u32, value: &[u8]) {
    put_key(output, field, 2);
    put_varint(output, value.len() as u64);
    output.extend_from_slice(value);
}

fn bytes_field_with_framing(
    output: &mut Vec<u8>,
    field: u32,
    value: &[u8],
    canonical_key: bool,
    canonical_length: bool,
) {
    if canonical_key {
        put_key(output, field, 2);
    } else {
        // Field one with wire type two is key 0x0a; this two-byte varint is
        // deliberately equivalent and non-canonical.
        output.extend_from_slice(&[0x8a, 0x00]);
    }
    if canonical_length {
        put_varint(output, value.len() as u64);
    } else {
        // TABLE_ID is exactly 32 bytes, represented here with a redundant
        // continuation byte in the length varint.
        output.extend_from_slice(&[0xa0, 0x00]);
    }
    output.extend_from_slice(value);
}

fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
    put_key(output, field, 0);
    put_varint(output, value);
}

fn put_key(output: &mut Vec<u8>, field: u32, wire_type: u8) {
    put_varint(output, (u64::from(field) << 3) | u64::from(wire_type));
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value as u8) & 0x7f;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}
