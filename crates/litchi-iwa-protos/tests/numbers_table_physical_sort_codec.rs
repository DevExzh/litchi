//! Focused direct coverage for the public physical table-sort wire seam.
//!
//! These fixtures are deliberately assembled as protobuf wire bytes.  The
//! generated schema is private to `litchi-iwa-protos`; callers of this module
//! only provide source bytes and consume borrowed snapshots or rewritten
//! bytes.

use litchi_iwa_protos::table_physical_sort_codec as codec;

const MAX_FIELDS: usize = 1 << 12;
const MAX_RECORDS: usize = 64;
const MAX_ELEMENTS: usize = 1 << 12;
const MAX_RECURSION: u32 = 32;
const MAX_WORK_BYTES: usize = 1 << 20;
const MAX_SCRATCH_BYTES: usize = 1 << 20;

fn options(source: &[u8]) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        source.len().max(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_RECORDS,
        MAX_ELEMENTS,
        MAX_WORK_BYTES,
        MAX_SCRATCH_BYTES,
    )
}

fn options_with(
    max_message_bytes: usize,
    max_work_bytes: usize,
    max_output_bytes: usize,
) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        max_message_bytes,
        MAX_FIELDS,
        max_work_bytes,
        MAX_RECURSION,
        MAX_RECORDS,
        MAX_ELEMENTS,
        max_output_bytes,
        MAX_SCRATCH_BYTES,
    )
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = u8::try_from(value & 0x7f).expect("seven bits fit in a byte");
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

fn field_varint(output: &mut Vec<u8>, field: u32, value: u64) {
    push_varint(output, u64::from(field) << 3);
    push_varint(output, value);
}

fn field_bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    push_varint(output, (u64::from(field) << 3) | 2);
    push_varint(
        output,
        u64::try_from(payload.len()).expect("test payload length fits in u64"),
    );
    output.extend_from_slice(payload);
}

fn field_fixed32(output: &mut Vec<u8>, field: u32, value: u32) {
    push_varint(output, (u64::from(field) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
}

fn group_field(field: u32, body: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(&mut output, (u64::from(field) << 3) | 3);
    output.extend_from_slice(body);
    push_varint(&mut output, (u64::from(field) << 3) | 4);
    output
}

fn tile_row_with_unknown(index: u32, marker: u8, include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, u64::from(index));
    field_varint(&mut output, 2, u64::from(index + 1));
    field_bytes(&mut output, 3, &[0xa0, marker]);
    field_bytes(&mut output, 4, &[0xb0, marker]);
    field_varint(&mut output, 5, 7);
    field_bytes(&mut output, 6, &[0xc0, marker]);
    field_bytes(&mut output, 7, &[0xd0, marker]);
    field_varint(&mut output, 8, u64::from(index & 1));
    if include_unknown {
        // The marker is inside an unknown nested field and must remain
        // source-preserved for a no-op read.
        field_bytes(&mut output, 90, &[0xe0, marker]);
    }
    output
}

fn tile_fixture() -> Vec<u8> {
    tile_fixture_with_unknowns(true)
}

fn tile_fixture_clean() -> Vec<u8> {
    tile_fixture_with_unknowns(false)
}

fn tile_fixture_with_unknowns(include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, 4);
    field_varint(&mut output, 2, 3);
    field_varint(&mut output, 3, 3);
    field_varint(&mut output, 4, 3);

    if include_unknown {
        let root_unknown = group_field(92, &[0x08, 0x01, 0x9a, 0x06, 0x02, 0xca, 0xfe]);
        field_bytes(&mut output, 90, b"tile-root-before");
        output.extend_from_slice(&root_unknown);
    }
    field_bytes(
        &mut output,
        5,
        &tile_row_with_unknown(0, b'a', include_unknown),
    );
    if include_unknown {
        field_bytes(&mut output, 91, b"tile-root-middle");
    }
    field_bytes(
        &mut output,
        5,
        &tile_row_with_unknown(1, b'b', include_unknown),
    );
    field_bytes(
        &mut output,
        5,
        &tile_row_with_unknown(2, b'c', include_unknown),
    );
    field_varint(&mut output, 6, 11);
    field_varint(&mut output, 7, 1);
    if include_unknown {
        output.extend_from_slice(&group_field(93, &[0x10, 0x02]));
    }
    output
}

fn header_record_with_unknown(index: u32, marker: u8, include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, u64::from(index));
    field_fixed32(&mut output, 2, (16.0_f32 + f32::from(marker)).to_bits());
    field_varint(&mut output, 3, u64::from(index & 1));
    field_varint(&mut output, 4, u64::from(index + 1));
    field_bytes(&mut output, 5, &[0x10, marker]);
    field_bytes(&mut output, 6, &[0x20, marker]);
    if include_unknown {
        // Unknown bytes remain source-preserved for a no-op read.
        field_bytes(&mut output, 77, &[0x70, marker]);
    }
    output
}

fn header_fixture() -> Vec<u8> {
    header_fixture_with_unknowns(true)
}

fn header_fixture_clean() -> Vec<u8> {
    header_fixture_with_unknowns(false)
}

fn header_fixture_with_unknowns(include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, 17);
    if include_unknown {
        let root_unknown = group_field(92, &[0x08, 0x05]);
        output.extend_from_slice(&root_unknown);
    }
    field_bytes(
        &mut output,
        2,
        &header_record_with_unknown(0, b'x', include_unknown),
    );
    if include_unknown {
        field_bytes(&mut output, 90, b"header-root-middle");
    }
    field_bytes(
        &mut output,
        2,
        &header_record_with_unknown(4, b'y', include_unknown),
    );
    field_bytes(
        &mut output,
        2,
        &header_record_with_unknown(9, b'z', include_unknown),
    );
    if include_unknown {
        field_varint(&mut output, 91, 23);
    }
    output
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, lower);
    field_varint(&mut output, 2, upper);
    output
}

fn uid_map_fixture() -> Vec<u8> {
    uid_map_fixture_with_unknowns(true)
}

fn uid_map_fixture_clean() -> Vec<u8> {
    uid_map_fixture_with_unknowns(false)
}

fn uid_map_fixture_with_unknowns(include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    for (lower, upper) in [(11, 101), (12, 102)] {
        field_bytes(&mut output, 1, &uuid(lower, upper));
    }
    for value in [0, 1] {
        field_varint(&mut output, 2, value);
    }
    for value in [0, 1] {
        field_varint(&mut output, 3, value);
    }
    for (lower, upper) in [(21, 201), (22, 202), (23, 203)] {
        field_bytes(&mut output, 4, &uuid(lower, upper));
    }
    for value in [0, 1, 2] {
        field_varint(&mut output, 5, value);
    }
    for value in [0, 1, 2] {
        field_varint(&mut output, 6, value);
    }
    if include_unknown {
        output.extend_from_slice(&group_field(90, &[0x08, 0x01, 0x10, 0x02]));
        field_bytes(&mut output, 91, b"uid-root-unknown");
    }
    output
}

fn contains_bytes(source: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty() && source.windows(needle.len()).any(|window| window == needle)
}

fn assert_borrowed(source: &[u8], borrowed: &[u8]) {
    assert!(!borrowed.is_empty());
    let source_start = source.as_ptr() as usize;
    let source_end = source_start + source.len();
    let borrowed_start = borrowed.as_ptr() as usize;
    let borrowed_end = borrowed_start + borrowed.len();
    assert!(borrowed_start >= source_start && borrowed_end <= source_end);
}

fn assert_limit(error: codec::DecodeError, expected: fn(codec::DecodeLimit) -> bool) {
    let limit = error.resource_limit().expect("limit should be reported");
    assert!(expected(limit));
}

fn is_input_limit(limit: codec::DecodeLimit) -> bool {
    matches!(limit, codec::DecodeLimit::Bytes { .. })
}

fn is_output_limit(limit: codec::DecodeLimit) -> bool {
    matches!(limit, codec::DecodeLimit::OutputBytes { .. })
}

fn is_work_limit(limit: codec::DecodeLimit) -> bool {
    matches!(limit, codec::DecodeLimit::Work { .. })
}

#[test]
fn tile_plan_borrows_rows_and_rejects_unknown_mutation() {
    let unknown_source = tile_fixture();
    let unknown_before = unknown_source.clone();
    let (unknown_identity, unknown_report) =
        codec::rewrite_tile_rows(&unknown_source, 8, &[], options(&unknown_source))
            .expect("unknown tile no-op executes");
    assert_eq!(unknown_identity, unknown_source);
    assert_eq!(unknown_report.changed_records(), 0);
    let unknown_moves = [codec::RowMove::new(0, 2), codec::RowMove::new(2, 0)];
    assert!(
        codec::rewrite_tile_rows(&unknown_source, 8, &unknown_moves, options(&unknown_source),)
            .is_err()
    );
    assert_eq!(unknown_source, unknown_before);

    let source = tile_fixture_clean();
    let before = source.clone();
    let moves = [codec::RowMove::new(0, 2), codec::RowMove::new(2, 0)];
    let plan = codec::plan_tile_rows_rewrite(&source, 8, &moves, options(&source))
        .expect("valid tile plan");

    assert_eq!(plan.tile().num_rows(), 3);
    let rows: Vec<_> = plan.row_records().collect();
    assert_eq!(rows.len(), 3);
    for row in &rows {
        assert_borrowed(&source, row.raw());
        let snapshot = row.snapshot();
        assert_borrowed(&source, snapshot.cell_storage_buffer_pre_bnc());
        assert_borrowed(&source, snapshot.cell_offsets_pre_bnc());
        assert_borrowed(
            &source,
            snapshot
                .cell_storage_buffer()
                .expect("fixture has current storage"),
        );
        assert_borrowed(
            &source,
            snapshot
                .cell_offsets()
                .expect("fixture has current offsets"),
        );
    }
    assert_eq!(rows[0].snapshot().tile_row_index(), 0);
    assert_eq!(rows[1].snapshot().tile_row_index(), 1);
    assert_eq!(rows[2].snapshot().tile_row_index(), 2);
    assert_eq!(
        source, before,
        "planning must not mutate its borrowed source"
    );

    let (candidate, report) =
        codec::execute_tile_rows_rewrite(plan, options(&source)).expect("valid tile plan executes");
    assert_eq!(
        source, before,
        "execution must not mutate its borrowed source"
    );
    assert_eq!(report.changed_records(), 2);
    assert_eq!(report.output_bytes(), candidate.len());

    let storage_version = {
        let mut bytes = Vec::new();
        field_varint(&mut bytes, 6, 11);
        bytes
    };
    assert!(contains_bytes(&candidate, &storage_version));

    let reparsed = codec::plan_tile_rows_rewrite(&candidate, 8, &[], options(&candidate))
        .expect("rewritten tile reparses");
    let reparsed_rows: Vec<_> = reparsed.row_records().collect();
    let markers: Vec<_> = reparsed_rows
        .iter()
        .map(|row| row.snapshot().cell_storage_buffer_pre_bnc()[1])
        .collect();
    let indexes: Vec<_> = reparsed_rows
        .iter()
        .map(|row| row.snapshot().tile_row_index())
        .collect();
    assert_eq!(indexes, [0, 1, 2]);
    assert_eq!(markers, [b'c', b'b', b'a']);
    for (marker, row) in [
        (b'a', &reparsed_rows[2]),
        (b'b', &reparsed_rows[1]),
        (b'c', &reparsed_rows[0]),
    ] {
        assert!(contains_bytes(
            row.raw(),
            &field_bytes_vec(3, &[0xa0, marker])
        ));
    }
    let (stable, _) = codec::execute_tile_rows_rewrite(reparsed, options(&candidate))
        .expect("rewritten tile no-op executes");
    assert_eq!(stable, candidate);
}

#[test]
fn sparse_header_moves_borrow_records_and_reject_unknown_mutation() {
    let unknown_source = header_fixture();
    let unknown_before = unknown_source.clone();
    let (unknown_identity, unknown_report) = codec::rewrite_header_storage_bucket_rows(
        &unknown_source,
        16,
        &[],
        options(&unknown_source),
    )
    .expect("unknown header no-op executes");
    assert_eq!(unknown_identity, unknown_source);
    assert_eq!(unknown_report.changed_records(), 0);
    let unknown_moves = [
        codec::HeaderRowMove::new(0, 9),
        codec::HeaderRowMove::new(9, 0),
    ];
    assert!(
        codec::rewrite_header_storage_bucket_rows(
            &unknown_source,
            16,
            &unknown_moves,
            options(&unknown_source),
        )
        .is_err()
    );
    assert_eq!(unknown_source, unknown_before);

    let source = header_fixture_clean();
    let before = source.clone();
    let moves = [
        codec::HeaderRowMove::new(0, 9),
        codec::HeaderRowMove::new(9, 0),
    ];
    let plan = codec::plan_header_storage_bucket_rows(&source, 16, &moves, options(&source))
        .expect("valid sparse header plan");
    let records: Vec<_> = plan.records().collect();
    assert_eq!(records.len(), 3);
    for record in &records {
        assert_borrowed(&source, record.raw());
    }
    assert_eq!(
        records
            .iter()
            .map(|record| record.snapshot().index())
            .collect::<Vec<_>>(),
        [0, 4, 9]
    );
    assert!(records[0].snapshot().has_cell_style());
    assert!(records[0].snapshot().has_text_style());
    assert_eq!(source, before, "header planning must not mutate source");

    let (candidate, report) = codec::execute_header_storage_bucket_rows(plan, options(&source))
        .expect("valid sparse header plan executes");
    assert_eq!(report.changed_records(), 2);
    assert_eq!(source, before, "header execution must not mutate source");
    assert!(contains_bytes(
        &candidate,
        &field_bytes_vec(5, &[0x10, b'x'])
    ));

    let reparsed = codec::plan_header_storage_bucket_rows(&candidate, 16, &[], options(&candidate))
        .expect("rewritten sparse headers reparse");
    let reparsed_records: Vec<_> = reparsed.records().collect();
    let indexes: Vec<_> = reparsed_records
        .iter()
        .map(|record| record.snapshot().index())
        .collect();
    let markers: Vec<_> = reparsed_records
        .iter()
        .map(|record| {
            record.raw().windows(4).find_map(|window| {
                (window[0] == 0x2a && window[1] == 2 && window[2] == 0x10).then_some(window[3])
            })
        })
        .collect();
    assert_eq!(indexes, [0, 4, 9]);
    assert_eq!(markers, [Some(b'z'), Some(b'y'), Some(b'x')]);
    let (stable, _) = codec::execute_header_storage_bucket_rows(reparsed, options(&candidate))
        .expect("rewritten sparse headers no-op executes");
    assert_eq!(stable, candidate);

    // Sorting the moves before validation is intentional, so duplicate source
    // indexes must be rejected even when they are separated in caller order.
    let nonadjacent_duplicate = [
        codec::HeaderRowMove::new(0, 9),
        codec::HeaderRowMove::new(4, 4),
        codec::HeaderRowMove::new(0, 0),
    ];
    assert!(
        codec::plan_header_storage_bucket_rows(
            &source,
            16,
            &nonadjacent_duplicate,
            options(&source),
        )
        .is_err()
    );
    // Header indexes are sparse; an absent source index is not an implicit
    // insertion and must fail before any rewrite is staged.
    let missing_source = [codec::HeaderRowMove::new(1, 9)];
    assert!(
        codec::plan_header_storage_bucket_rows(&source, 16, &missing_source, options(&source),)
            .is_err()
    );
    assert_eq!(source, before);
}

#[test]
fn uid_map_plan_rewrites_both_inverse_arrays_and_rejects_unknown_mutation() {
    let unknown_source = uid_map_fixture();
    let unknown_before = unknown_source.clone();
    let identity = codec::RowUidPermutation::identity(3).expect("identity UID permutation");
    let (unknown_identity, unknown_report) = codec::rewrite_column_row_uid_map(
        &unknown_source,
        2,
        3,
        &identity,
        options(&unknown_source),
    )
    .expect("unknown UID map no-op executes");
    assert_eq!(unknown_identity, unknown_source);
    assert_eq!(unknown_report.changed_records(), 0);
    let unknown_permutation =
        codec::RowUidPermutation::new(&[2, 0, 1]).expect("valid unknown UID permutation");
    assert!(
        codec::rewrite_column_row_uid_map(
            &unknown_source,
            2,
            3,
            &unknown_permutation,
            options(&unknown_source),
        )
        .is_err()
    );
    assert_eq!(unknown_source, unknown_before);

    let source = uid_map_fixture_clean();
    let before = source.clone();
    let permutation = codec::RowUidPermutation::new(&[2, 0, 1]).expect("valid permutation");
    let plan =
        codec::plan_column_row_uid_map_rewrite(&source, 2, 3, &permutation, options(&source))
            .expect("valid UID-map plan");
    assert_eq!(plan.row_uid_for_index(), [1, 2, 0]);
    assert_eq!(plan.row_index_for_uid(), [2, 0, 1]);
    assert_eq!(source, before, "UID planning must not mutate source");

    let (candidate, report) =
        codec::execute_column_row_uid_map_rewrite(plan, 2, 3, options(&source))
            .expect("valid UID-map plan executes");
    assert!(report.changed_records() > 0);
    assert_eq!(source, before, "UID execution must not mutate source");
    let (snapshot, _) = codec::decode_column_row_uid_map(&candidate, 2, 3, options(&candidate))
        .expect("rewritten UID map reparses");
    assert_eq!(snapshot.row_uid_for_index(), [1, 2, 0]);
    assert_eq!(snapshot.row_index_for_uid(), [2, 0, 1]);
}

#[test]
fn malformed_and_noncanonical_wire_is_rejected_at_public_boundary() {
    let tile_truncated = [0x08, 0x01, 0x10];
    assert!(
        codec::plan_tile_rows_rewrite(&tile_truncated, 8, &[], options(&tile_truncated)).is_err()
    );

    let mut tile_noncanonical = tile_fixture();
    // Field 90 is unknown, but its key and value still have to be canonical.
    tile_noncanonical.extend_from_slice(&[0xd0, 0x05, 0x80, 0x00]);
    assert!(
        codec::plan_tile_rows_rewrite(&tile_noncanonical, 8, &[], options(&tile_noncanonical),)
            .is_err()
    );

    let header_wrong_wire = [0x08, 0x11, 0x12, 0x04, 0x08, 0x00, 0x18, 0x00];
    assert!(
        codec::plan_header_storage_bucket_rows(
            &header_wrong_wire,
            16,
            &[],
            options(&header_wrong_wire),
        )
        .is_err()
    );

    let mut uid_noncanonical = uid_map_fixture();
    uid_noncanonical.extend_from_slice(&[0x48, 0x80, 0x00]);
    assert!(
        codec::decode_column_row_uid_map(&uid_noncanonical, 2, 3, options(&uid_noncanonical))
            .is_err()
    );

    let uid_truncated = [0x0a, 0x02, 0x08];
    assert!(
        codec::decode_column_row_uid_map(&uid_truncated, 1, 1, options(&uid_truncated)).is_err()
    );
}

#[test]
fn one_below_input_output_and_work_limits_refuse_tile_plan_before_publication() {
    let source = tile_fixture_clean();
    let moves = [codec::RowMove::new(0, 2), codec::RowMove::new(2, 0)];
    let requirements = codec::plan_tile_rows_rewrite(&source, 8, &moves, options(&source))
        .expect("valid tile plan")
        .requirements();

    let input_options = options_with(source.len() - 1, MAX_WORK_BYTES, MAX_WORK_BYTES);
    let input_error = codec::plan_tile_rows_rewrite(&source, 8, &moves, input_options)
        .expect_err("one-below input budget must reject");
    assert_limit(input_error, is_input_limit);

    let output_options = options_with(
        source.len(),
        MAX_WORK_BYTES,
        requirements.output_bytes() - 1,
    );
    let output_error = codec::plan_tile_rows_rewrite(&source, 8, &moves, output_options)
        .expect_err("one-below output budget must reject");
    assert_limit(output_error, is_output_limit);

    let work_options = options_with(source.len(), requirements.work_bytes() - 1, MAX_WORK_BYTES);
    let work_error = codec::plan_tile_rows_rewrite(&source, 8, &moves, work_options)
        .expect_err("one-below work budget must reject");
    assert_limit(work_error, is_work_limit);

    // A prepared plan applies the same output/work checks before allocating a
    // candidate, so the caller can safely retain the source until publication.
    let output_plan = codec::plan_tile_rows_rewrite(&source, 8, &moves, options(&source))
        .expect("valid tile plan");
    let output_error = codec::execute_tile_rows_rewrite(output_plan, output_options)
        .expect_err("one-below output budget must reject execution");
    assert_limit(output_error, is_output_limit);
    let work_plan = codec::plan_tile_rows_rewrite(&source, 8, &moves, options(&source))
        .expect("valid tile plan");
    let work_error = codec::execute_tile_rows_rewrite(work_plan, work_options)
        .expect_err("one-below work budget must reject execution");
    assert_limit(work_error, is_work_limit);
}

fn field_bytes_vec(field: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    field_bytes(&mut output, field, payload);
    output
}
