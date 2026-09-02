#![no_main]

//! Bounded fuzzing for the source-preserving physical table-sort wire seam.
//!
//! The target deliberately exercises the three independent storage envelopes
//! used by the Keynote adapter: tile row records, sparse row-header records,
//! and the stable row/column UID map.  A successful plan is executed twice
//! (the prepared and one-shot routes), and the source is checked after every
//! call.  Fixed recipes keep strict malformed, duplicate, non-canonical,
//! truncated, unknown-field, and exact-resource-boundary paths reachable even
//! when libFuzzer starts with arbitrary bytes.

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_protos::keynote_table_physical_sort_codec as codec;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_RECORDS: usize = 4 * 1024;
const MAX_ELEMENTS: usize = 16 * 1024;
const MAX_SCRATCH_BYTES: usize = 512 * 1024;
const MAX_RECURSION: u32 = 64;

const TILE_ROOT_UNKNOWN: &[u8] = b"physical-sort-tile-root-unknown";
const HEADER_ROOT_UNKNOWN: &[u8] = b"physical-sort-header-root-unknown";
const UNKNOWN_GROUP: &[u8] = &[0xdb, 0x05, 0x08, 0x01, 0xdc, 0x05];

fuzz_target!(|data: &[u8]| {
    let Some(source) = normalize_input(data) else {
        return;
    };

    exercise_tile(&source, data);
    exercise_headers(&source, data);
    exercise_uid_map(&source, data);

    // Deterministic fixtures are run once per worker.  This keeps valid
    // planned rewrites, unknown-mutation refusals, and every malformed family
    // hot without adding native package bytes to the checked-in corpus.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(|| {
        let tiles = tile_fixtures();
        for source in &tiles {
            exercise_tile(source, b"fixed-keynote-physical-tile");
        }
        let headers = header_fixtures();
        for source in &headers {
            exercise_headers(source, b"fixed-keynote-physical-header");
        }
        let uid_maps = uid_fixtures();
        for source in &uid_maps {
            exercise_uid_map(source, b"fixed-keynote-physical-uid");
        }
        assert_unknown_mutations_reject_atomically(&tiles[0], &headers[0], &uid_maps[0]);
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
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECURSION,
        MAX_RECORDS,
        MAX_ELEMENTS,
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_SCRATCH_BYTES,
    )
}

fn options_with(
    _source: &[u8],
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_records: usize,
    max_elements: usize,
    max_output_bytes: usize,
    max_scratch_bytes: usize,
) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        max_message_bytes,
        max_fields,
        max_work_bytes,
        MAX_RECURSION,
        max_records,
        max_elements,
        max_output_bytes,
        max_scratch_bytes,
    )
}

fn assert_unknown_mutations_reject_atomically(tile: &[u8], header: &[u8], uid_map: &[u8]) {
    let tile_before = tile.to_vec();
    let tile_moves = [codec::RowMove::new(0, 1), codec::RowMove::new(1, 0)];
    assert!(codec::plan_tile_rows_rewrite(tile, u32::MAX, &tile_moves, options(tile)).is_err());
    assert!(codec::rewrite_tile_rows(tile, u32::MAX, &tile_moves, options(tile)).is_err());
    assert_eq!(
        tile,
        tile_before.as_slice(),
        "tile refusal modified its source"
    );

    let header_before = header.to_vec();
    let header_moves = [
        codec::HeaderRowMove::new(0, 2),
        codec::HeaderRowMove::new(2, 0),
    ];
    assert!(
        codec::plan_header_storage_bucket_rows(header, u32::MAX, &header_moves, options(header),)
            .is_err()
    );
    assert!(codec::rewrite_header_storage_bucket_rows(
        header,
        u32::MAX,
        &header_moves,
        options(header),
    )
    .is_err());
    assert_eq!(
        header,
        header_before.as_slice(),
        "header refusal modified source"
    );

    let uid_before = uid_map.to_vec();
    let permutation = codec::RowUidPermutation::new(&[1, 0, 2])
        .unwrap_or_else(|error| panic!("unknown UID fixture permutation rejected: {error}"));
    assert!(
        codec::plan_column_row_uid_map_rewrite(uid_map, 2, 3, &permutation, options(uid_map),)
            .is_err()
    );
    assert!(
        codec::rewrite_column_row_uid_map(uid_map, 2, 3, &permutation, options(uid_map),).is_err()
    );
    assert_eq!(
        uid_map,
        uid_before.as_slice(),
        "UID refusal modified source"
    );
}

fn exercise_tile(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let generous = options(source);

    // An empty move list is a strict decode, source-borrowing plan, and
    // source-preserving no-op.  It is also the scalar/readback half of the
    // prepared-vs-one-shot differential check.
    let parsed = codec::plan_tile_rows_rewrite(source, u32::MAX, &[], generous);
    assert_eq!(
        source,
        before.as_slice(),
        "tile planning modified its source"
    );
    let plan = match parsed {
        Ok(plan) => plan,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let tile = plan.tile();
    let rows: Vec<_> = plan.row_records().collect();
    assert!(rows.len() <= MAX_RECORDS);
    for row in &rows {
        assert_borrowed(source, row.raw());
        black_box(row.snapshot());
    }
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    let (no_op, no_op_report) = codec::execute_tile_rows_rewrite(plan, generous)
        .unwrap_or_else(|error| panic!("valid tile no-op plan failed to execute: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "tile execution modified its source"
    );
    assert_eq!(no_op, source, "tile no-op changed source bytes");
    assert_eq!(no_op_report.output_bytes(), no_op.len());
    assert_report(no_op_report.source(), source.len());
    assert_report(no_op_report.result(), no_op.len());

    let one_shot = codec::rewrite_tile_rows(source, tile.num_rows(), &[], generous);
    assert_eq!(
        source,
        before.as_slice(),
        "tile one-shot modified its source"
    );
    if let Ok((candidate, report)) = one_shot {
        assert_eq!(
            candidate, source,
            "tile one-shot no-op changed source bytes"
        );
        assert_eq!(report, no_op_report);
    }

    if tile.num_rows() < 2 {
        exercise_tile_limits(source, tile.num_rows(), &[], generous);
        return;
    }

    let moves = [codec::RowMove::new(0, 1), codec::RowMove::new(1, 0)];
    let planned = codec::plan_tile_rows_rewrite(source, tile.num_rows(), &moves, generous);
    assert_eq!(
        source,
        before.as_slice(),
        "tile rewrite planning modified source"
    );
    let plan = match planned {
        Ok(plan) => plan,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    let (candidate, report) = codec::execute_tile_rows_rewrite(plan, generous)
        .unwrap_or_else(|error| panic!("valid tile rewrite plan failed to execute: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "tile rewrite modified its source"
    );
    assert_eq!(candidate.len(), requirements.output_bytes());
    assert_eq!(report.output_bytes(), candidate.len());
    assert_eq!(report.changed_records(), 2);
    assert_report(report.source(), source.len());
    assert_report(report.result(), candidate.len());

    // The one-shot wrapper must use the same preflight and preserve the same
    // known source spans as the prepared plan.  Unknown mutable envelopes
    // return early above and are covered by the deterministic refusal check.
    let one_shot = codec::rewrite_tile_rows(source, tile.num_rows(), &moves, generous)
        .unwrap_or_else(|error| panic!("valid tile one-shot rewrite failed: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "tile one-shot rewrite modified source"
    );
    assert_eq!(one_shot.0, candidate);
    assert_eq!(one_shot.1, report);
    // Reparse the candidate through a strict no-op and require exact byte
    // stability.  This catches output records that are valid in isolation but
    // no longer satisfy the enclosing Tile invariants.
    let candidate_options = options(&candidate);
    let candidate_plan =
        codec::plan_tile_rows_rewrite(&candidate, u32::MAX, &[], candidate_options)
            .unwrap_or_else(|error| panic!("rewritten tile failed strict reparse: {error}"));
    let (reparsed, _) = codec::execute_tile_rows_rewrite(candidate_plan, candidate_options)
        .unwrap_or_else(|error| panic!("rewritten tile no-op failed: {error}"));
    assert_eq!(reparsed, candidate);

    exercise_tile_limits(source, tile.num_rows(), &moves, generous);
    black_box(data);
}

fn exercise_tile_limits(
    source: &[u8],
    row_index_limit: u32,
    moves: &[codec::RowMove],
    generous: codec::DecodeOptions,
) {
    let plan = match codec::plan_tile_rows_rewrite(source, row_index_limit, moves, generous) {
        Ok(plan) => plan,
        Err(_) => return,
    };
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());

    // Rebuild the plan for each profile because execution consumes it.  Each
    // profile is one byte below exactly one preflight requirement, so failure
    // must occur before candidate publication/allocation.
    if requirements.output_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECORDS,
            MAX_ELEMENTS,
            requirements.output_bytes() - 1,
            MAX_SCRATCH_BYTES,
        );
        let result = codec::plan_tile_rows_rewrite(source, row_index_limit, moves, constrained);
        assert!(result.is_err(), "one-below tile output budget was accepted");
        observe_error(result.err().expect("one-below tile output error"));
    }
    if requirements.work_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            requirements.work_bytes() - 1,
            MAX_RECORDS,
            MAX_ELEMENTS,
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_SCRATCH_BYTES,
        );
        let result = codec::plan_tile_rows_rewrite(source, row_index_limit, moves, constrained);
        assert!(result.is_err(), "one-below tile work budget was accepted");
        observe_error(result.err().expect("one-below tile work error"));
    }
    let constrained = options_with(
        source,
        source.len().saturating_sub(1),
        MAX_FIELDS,
        MAX_WORK_BYTES,
        MAX_RECORDS,
        MAX_ELEMENTS,
        MAX_OUTPUT_BYTES.max(source.len()),
        MAX_SCRATCH_BYTES,
    );
    let result = codec::plan_tile_rows_rewrite(source, row_index_limit, moves, constrained);
    assert!(result.is_err(), "one-below tile input budget was accepted");
    observe_error(result.err().expect("one-below tile input error"));
}

fn exercise_headers(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let generous = options(source);
    let parsed = codec::plan_header_storage_bucket_rows(source, u32::MAX, &[], generous);
    assert_eq!(
        source,
        before.as_slice(),
        "header planning modified its source"
    );
    let plan = match parsed {
        Ok(plan) => plan,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert!(plan.records().count() <= MAX_RECORDS);
    for record in plan.records() {
        assert_borrowed(source, record.raw());
        black_box(record.snapshot());
    }
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    let (no_op, no_op_report) = codec::execute_header_storage_bucket_rows(plan, generous)
        .unwrap_or_else(|error| panic!("valid header no-op plan failed: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "header execution modified its source"
    );
    assert_eq!(no_op, source, "header no-op changed source bytes");
    assert_eq!(no_op_report.output_bytes(), no_op.len());
    assert_report(no_op_report.source(), source.len());
    assert_report(no_op_report.result(), no_op.len());
    let one_shot = codec::rewrite_header_storage_bucket_rows(source, u32::MAX, &[], generous);
    assert_eq!(
        source,
        before.as_slice(),
        "header one-shot modified its source"
    );
    if let Ok((candidate, report)) = one_shot {
        assert_eq!(candidate, source, "header one-shot no-op changed bytes");
        assert_eq!(report, no_op_report);
    }

    let Some((first, second)) = first_two_header_indexes(source) else {
        return;
    };
    let moves = [
        codec::HeaderRowMove::new(first, second),
        codec::HeaderRowMove::new(second, first),
    ];
    let planned = codec::plan_header_storage_bucket_rows(source, u32::MAX, &moves, generous);
    assert_eq!(
        source,
        before.as_slice(),
        "header rewrite planning modified source"
    );
    let plan = match planned {
        Ok(plan) => plan,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    let (candidate, report) = codec::execute_header_storage_bucket_rows(plan, generous)
        .unwrap_or_else(|error| panic!("valid header rewrite plan failed: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "header rewrite modified its source"
    );
    assert_eq!(candidate.len(), requirements.output_bytes());
    assert_eq!(report.output_bytes(), candidate.len());
    assert_eq!(report.changed_records(), 2);
    let one_shot = codec::rewrite_header_storage_bucket_rows(source, u32::MAX, &moves, generous)
        .unwrap_or_else(|error| panic!("valid header one-shot rewrite failed: {error}"));
    assert_eq!(
        source,
        before.as_slice(),
        "header one-shot rewrite modified source"
    );
    assert_eq!(one_shot.0, candidate);
    assert_eq!(one_shot.1, report);
    let candidate_options = options(&candidate);
    let candidate_plan =
        codec::plan_header_storage_bucket_rows(&candidate, u32::MAX, &[], candidate_options)
            .unwrap_or_else(|error| panic!("rewritten header failed strict reparse: {error}"));
    let (reparsed, _) =
        codec::execute_header_storage_bucket_rows(candidate_plan, candidate_options)
            .unwrap_or_else(|error| panic!("rewritten header no-op failed: {error}"));
    assert_eq!(reparsed, candidate);
    exercise_header_limits(source, &moves);
    black_box(data);
}

fn exercise_header_limits(source: &[u8], moves: &[codec::HeaderRowMove]) {
    let generous = options(source);
    let plan = match codec::plan_header_storage_bucket_rows(source, u32::MAX, moves, generous) {
        Ok(plan) => plan,
        Err(_) => return,
    };
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    if requirements.output_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECORDS,
            MAX_ELEMENTS,
            requirements.output_bytes() - 1,
            MAX_SCRATCH_BYTES,
        );
        let result = codec::plan_header_storage_bucket_rows(source, u32::MAX, moves, constrained);
        assert!(
            result.is_err(),
            "one-below header output budget was accepted"
        );
        observe_error(result.err().expect("one-below header output error"));
    }
    if requirements.work_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            requirements.work_bytes() - 1,
            MAX_RECORDS,
            MAX_ELEMENTS,
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_SCRATCH_BYTES,
        );
        let result = codec::plan_header_storage_bucket_rows(source, u32::MAX, moves, constrained);
        assert!(result.is_err(), "one-below header work budget was accepted");
        observe_error(result.err().expect("one-below header work error"));
    }
}

fn exercise_uid_map(source: &[u8], data: &[u8]) {
    let before = source.to_vec();
    let generous = options(source);
    // A small dimension matrix makes valid mutated maps reachable without
    // allowing attacker-controlled dimensions to drive an allocation.
    for &(columns, rows) in &[(1, 1), (2, 2), (2, 3), (3, 3), (4, 4)] {
        let result = codec::decode_column_row_uid_map(source, columns, rows, generous);
        assert_eq!(source, before.as_slice(), "UID decode modified its source");
        let Ok((snapshot, report)) = result else {
            continue;
        };
        assert_eq!(snapshot.column_count(), columns);
        assert_eq!(snapshot.row_count(), rows);
        assert_eq!(snapshot.sorted_column_uids().len(), columns);
        assert_eq!(snapshot.column_index_for_uid().len(), columns);
        assert_eq!(snapshot.column_uid_for_index().len(), columns);
        assert_eq!(snapshot.sorted_row_uids().len(), rows);
        assert_eq!(snapshot.row_index_for_uid().len(), rows);
        assert_eq!(snapshot.row_uid_for_index().len(), rows);
        assert_report(report, source.len());
        assert_inverse(
            snapshot.column_index_for_uid(),
            snapshot.column_uid_for_index(),
        );
        assert_inverse(snapshot.row_index_for_uid(), snapshot.row_uid_for_index());
        let mut destination_by_source: Vec<u32> = (0..rows as u32).collect();
        if destination_by_source.len() > 1 {
            destination_by_source.rotate_left(1);
        }
        let permutation = codec::RowUidPermutation::new(&destination_by_source)
            .unwrap_or_else(|error| panic!("generated UID permutation rejected: {error}"));
        let planned =
            codec::plan_column_row_uid_map_rewrite(source, columns, rows, &permutation, generous);
        assert_eq!(
            source,
            before.as_slice(),
            "UID planning modified its source"
        );
        let plan = match planned {
            Ok(plan) => plan,
            Err(error) => {
                observe_error(error);
                continue;
            },
        };
        let requirements = plan.requirements();
        assert_requirements(requirements, source.len());
        let (candidate, report) =
            codec::execute_column_row_uid_map_rewrite(plan, columns, rows, generous)
                .unwrap_or_else(|error| panic!("valid UID rewrite plan failed: {error}"));
        assert_eq!(source, before.as_slice(), "UID rewrite modified its source");
        assert_eq!(candidate.len(), requirements.output_bytes());
        assert_eq!(report.output_bytes(), candidate.len());
        assert_report(report.source(), source.len());
        assert_report(report.result(), candidate.len());
        let one_shot =
            codec::rewrite_column_row_uid_map(source, columns, rows, &permutation, generous)
                .unwrap_or_else(|error| panic!("valid UID one-shot rewrite failed: {error}"));
        assert_eq!(
            source,
            before.as_slice(),
            "UID one-shot modified its source"
        );
        assert_eq!(one_shot.0, candidate);
        assert_eq!(one_shot.1, report);
        let candidate_result =
            codec::decode_column_row_uid_map(&candidate, columns, rows, options(&candidate))
                .unwrap_or_else(|error| panic!("rewritten UID map failed strict reparse: {error}"));
        assert_eq!(candidate_result.0.column_count(), columns);
        assert_eq!(candidate_result.0.row_count(), rows);
        exercise_uid_limits(source, columns, rows, &permutation);
        black_box(snapshot);
    }

    let count = usize::from(data.first().copied().unwrap_or_default() % 16);
    let mut permutation: Vec<u32> = (0..count as u32).collect();
    if count > 1 {
        permutation.rotate_left(1);
    }
    let valid = codec::RowUidPermutation::new(&permutation)
        .unwrap_or_else(|error| panic!("generated UID permutation rejected: {error}"));
    assert_eq!(valid.destination_by_source(), permutation.as_slice());
    assert_eq!(
        valid.destination_for_source(0),
        permutation.first().copied()
    );
    assert!(codec::RowUidPermutation::new(&[0, 0]).is_err());
    assert!(codec::RowUidPermutation::new(&[0, 2]).is_err());
    let identity = codec::RowUidPermutation::identity(count as u32)
        .unwrap_or_else(|error| panic!("identity UID permutation rejected: {error}"));
    assert_eq!(
        identity.destination_by_source(),
        permutation_identity(count).as_slice()
    );
}

fn exercise_uid_limits(
    source: &[u8],
    columns: usize,
    rows: usize,
    permutation: &codec::RowUidPermutation,
) {
    let generous = options(source);
    let plan = match codec::plan_column_row_uid_map_rewrite(
        source,
        columns,
        rows,
        permutation,
        generous,
    ) {
        Ok(plan) => plan,
        Err(_) => return,
    };
    let requirements = plan.requirements();
    assert_requirements(requirements, source.len());
    if requirements.output_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECORDS,
            MAX_ELEMENTS,
            requirements.output_bytes() - 1,
            MAX_SCRATCH_BYTES,
        );
        let result =
            codec::plan_column_row_uid_map_rewrite(source, columns, rows, permutation, constrained);
        assert!(result.is_err(), "one-below UID output budget was accepted");
        observe_error(result.err().expect("one-below UID output error"));
    }
    if requirements.work_bytes() > 0 {
        let constrained = options_with(
            source,
            MAX_INPUT_BYTES,
            MAX_FIELDS,
            requirements.work_bytes() - 1,
            MAX_RECORDS,
            MAX_ELEMENTS,
            MAX_OUTPUT_BYTES.max(source.len()),
            MAX_SCRATCH_BYTES,
        );
        let result =
            codec::plan_column_row_uid_map_rewrite(source, columns, rows, permutation, constrained);
        assert!(result.is_err(), "one-below UID work budget was accepted");
        observe_error(result.err().expect("one-below UID work error"));
    }
}

fn assert_inverse(index_for_uid: &[u32], uid_for_index: &[u32]) {
    for (uid, &index) in index_for_uid.iter().enumerate() {
        assert_eq!(uid_for_index.get(index as usize).copied(), Some(uid as u32));
    }
    for (index, &uid) in uid_for_index.iter().enumerate() {
        assert_eq!(index_for_uid.get(uid as usize).copied(), Some(index as u32));
    }
}

fn assert_requirements(requirements: codec::RewriteRequirements, source_len: usize) {
    let source = requirements.source();
    assert_eq!(source.source_bytes(), source_len);
    assert!(source.fields() <= MAX_FIELDS);
    assert!(source.work_bytes() <= MAX_WORK_BYTES);
    assert!(source.max_depth() <= MAX_RECURSION);
    assert!(source.records() <= MAX_RECORDS);
    assert!(source.elements() <= MAX_ELEMENTS);
    assert!(requirements.output_bytes() <= MAX_OUTPUT_BYTES.max(source_len));
    assert!(requirements.work_bytes() <= MAX_WORK_BYTES);
    assert!(requirements.fields() <= MAX_FIELDS);
    assert!(requirements.records() <= MAX_RECORDS);
    assert!(requirements.elements() <= MAX_ELEMENTS);
}

fn assert_report(report: codec::DecodeReport, source_len: usize) {
    assert_eq!(report.source_bytes(), source_len);
    assert!(report.fields() <= MAX_FIELDS);
    assert!(report.work_bytes() <= MAX_WORK_BYTES);
    assert!(report.max_depth() <= MAX_RECURSION);
    assert!(report.records() <= MAX_RECORDS);
    assert!(report.elements() <= MAX_ELEMENTS);
}

fn observe_error(error: codec::DecodeError) {
    if let Some(limit) = error.resource_limit() {
        match limit {
            codec::DecodeLimit::Bytes { observed, maximum }
            | codec::DecodeLimit::Fields { observed, maximum }
            | codec::DecodeLimit::Work { observed, maximum }
            | codec::DecodeLimit::Records { observed, maximum }
            | codec::DecodeLimit::Elements { observed, maximum }
            | codec::DecodeLimit::OutputBytes { observed, maximum }
            | codec::DecodeLimit::ScratchBytes { observed, maximum } => {
                black_box((observed, maximum));
            },
            codec::DecodeLimit::Nesting { observed, maximum } => {
                black_box((observed, maximum));
            },
            codec::DecodeLimit::Allocation { requested } => {
                black_box(requested);
            },
            _ => {
                black_box(limit);
            },
        }
    }
    black_box(error);
}

fn assert_borrowed(source: &[u8], payload: &[u8]) {
    if payload.is_empty() {
        return;
    }
    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("bounded source pointer range");
    let payload_start = payload.as_ptr() as usize;
    let payload_end = payload_start
        .checked_add(payload.len())
        .expect("bounded payload pointer range");
    assert!(
        payload_start >= source_start && payload_end <= source_end,
        "physical-sort record did not borrow from its source"
    );
}

fn first_two_header_indexes(source: &[u8]) -> Option<(u32, u32)> {
    // Fixed fixtures use a canonical HeaderStorageBucket.  This tiny scanner
    // intentionally recognizes only the repeated header envelope and returns
    // None for arbitrary bytes; strict codec admission remains authoritative.
    let mut indexes = [0u32; 2];
    let mut count = 0;
    let mut remaining = source;
    while let Some((field, rest)) = take_field(remaining) {
        remaining = rest;
        if field.number != 2 || field.wire != 2 {
            continue;
        }
        let mut payload = field.payload;
        while let Some((nested, nested_rest)) = take_field(payload) {
            payload = nested_rest;
            if nested.number == 1 && nested.wire == 0 {
                let Some(index) = decode_varint(nested.payload) else {
                    return None;
                };
                if count < indexes.len() {
                    indexes[count] = index as u32;
                    count += 1;
                }
                break;
            }
        }
        if count == indexes.len() {
            return Some((indexes[0], indexes[1]));
        }
    }
    None
}

#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: u8,
    payload: &'a [u8],
}

fn take_field(source: &[u8]) -> Option<(Field<'_>, &[u8])> {
    let (tag, tag_len) = decode_varint_with_len(source)?;
    let number = u32::try_from(tag >> 3).ok()?;
    let wire = u8::try_from(tag & 7).ok()?;
    if number == 0 {
        return None;
    }
    let rest = &source[tag_len..];
    match wire {
        0 => {
            let (_, value_len) = decode_varint_with_len(rest)?;
            Some((
                Field {
                    number,
                    wire,
                    payload: &rest[..value_len],
                },
                &rest[value_len..],
            ))
        },
        1 => Some((
            Field {
                number,
                wire,
                payload: rest.get(..8)?,
            },
            rest.get(8..)?,
        )),
        2 => {
            let (length, length_len) = decode_varint_with_len(rest)?;
            let length = usize::try_from(length).ok()?;
            let start = length_len;
            let end = start.checked_add(length)?;
            Some((
                Field {
                    number,
                    wire,
                    payload: rest.get(start..end)?,
                },
                rest.get(end..)?,
            ))
        },
        5 => Some((
            Field {
                number,
                wire,
                payload: rest.get(..4)?,
            },
            rest.get(4..)?,
        )),
        _ => None,
    }
}

fn decode_varint(source: &[u8]) -> Option<u64> {
    decode_varint_with_len(source).map(|(value, _)| value)
}

fn decode_varint_with_len(source: &[u8]) -> Option<(u64, usize)> {
    let mut value = 0u64;
    for (offset, byte) in source.iter().copied().enumerate().take(10) {
        let shift = offset * 7;
        value |= u64::from(byte & 0x7f).checked_shl(shift as u32)?;
        if byte & 0x80 == 0 {
            return Some((value, offset + 1));
        }
    }
    None
}

fn permutation_identity(count: usize) -> Vec<u32> {
    (0..count as u32).collect()
}

fn push_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    push_varint(output, u64::from(number) << 3);
    push_varint(output, value);
}

fn push_bytes_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) {
    push_varint(output, (u64::from(number) << 3) | 2);
    push_varint(output, payload.len() as u64);
    output.extend_from_slice(payload);
}

fn push_fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) {
    push_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_varint(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn tile_row(index: u32, marker: &[u8], include_unknown: bool) -> Vec<u8> {
    let mut row = Vec::new();
    push_varint_field(&mut row, 1, u64::from(index));
    push_varint_field(&mut row, 2, 1);
    push_bytes_field(&mut row, 3, &[0x11, index as u8]);
    push_bytes_field(&mut row, 4, &[0]);
    if include_unknown {
        push_bytes_field(&mut row, 90, marker);
    }
    row
}

fn tile_fixtures() -> Vec<Vec<u8>> {
    let mut valid = Vec::new();
    push_varint_field(&mut valid, 1, 2);
    push_varint_field(&mut valid, 2, 2);
    push_varint_field(&mut valid, 3, 3);
    push_varint_field(&mut valid, 4, 3);
    push_bytes_field(&mut valid, 90, TILE_ROOT_UNKNOWN);
    valid.extend_from_slice(UNKNOWN_GROUP);
    for (index, marker) in [
        (0, b"tile-row-0-unknown"),
        (1, b"tile-row-1-unknown"),
        (2, b"tile-row-2-unknown"),
    ] {
        push_bytes_field(&mut valid, 5, &tile_row(index, marker, true));
    }
    push_varint_field(&mut valid, 6, 5);

    let mut clean = Vec::new();
    push_varint_field(&mut clean, 1, 2);
    push_varint_field(&mut clean, 2, 2);
    push_varint_field(&mut clean, 3, 3);
    push_varint_field(&mut clean, 4, 3);
    for (index, marker) in [
        (0, b"tile-row-0-clean"),
        (1, b"tile-row-1-clean"),
        (2, b"tile-row-2-clean"),
    ] {
        push_bytes_field(&mut clean, 5, &tile_row(index, marker, false));
    }
    push_varint_field(&mut clean, 6, 5);

    let mut duplicate_root = valid.clone();
    push_varint_field(&mut duplicate_root, 1, 2);
    let mut duplicate_row = valid.clone();
    let duplicate = tile_row(0, b"duplicate-index", true);
    push_bytes_field(&mut duplicate_row, 5, &duplicate);
    let mut wrong_wire = vec![0x0d, 0x01, 0x00, 0x00, 0x00];
    wrong_wire.extend_from_slice(&valid[5..]);
    let noncanonical = vec![0x08, 0x82, 0x00, 0x10, 0x82, 0x00, 0x18, 0x01, 0x20, 0x01];
    let truncated = vec![
        0x08, 0x02, 0x10, 0x02, 0x18, 0x03, 0x20, 0x03, 0x2a, 0x05, 0x08,
    ];
    let unterminated_group = [
        0x08, 0x02, 0x10, 0x02, 0x18, 0x03, 0x20, 0x03, 0xdb, 0x05, 0x08, 0x01,
    ];
    let mismatched_group = [
        0x08, 0x02, 0x10, 0x02, 0x18, 0x03, 0x20, 0x03, 0xdb, 0x05, 0x08, 0x01, 0xe4, 0x05,
    ];
    vec![
        valid,
        clean,
        duplicate_root,
        duplicate_row,
        wrong_wire,
        noncanonical,
        truncated,
        unterminated_group.to_vec(),
        mismatched_group.to_vec(),
    ]
}

fn header_record(index: u32, marker: &[u8], include_unknown: bool) -> Vec<u8> {
    let mut header = Vec::new();
    push_varint_field(&mut header, 1, u64::from(index));
    push_fixed32_field(&mut header, 2, 32.0f32.to_bits());
    push_varint_field(&mut header, 3, 0);
    push_varint_field(&mut header, 4, 1);
    if include_unknown {
        push_bytes_field(&mut header, 90, marker);
    }
    header
}

fn header_fixtures() -> Vec<Vec<u8>> {
    let mut valid = Vec::new();
    push_varint_field(&mut valid, 1, 7);
    push_bytes_field(&mut valid, 90, HEADER_ROOT_UNKNOWN);
    valid.extend_from_slice(UNKNOWN_GROUP);
    push_bytes_field(&mut valid, 2, &header_record(0, b"header-0-unknown", true));
    push_bytes_field(&mut valid, 2, &header_record(2, b"header-2-unknown", true));

    let mut clean = Vec::new();
    push_varint_field(&mut clean, 1, 7);
    push_bytes_field(&mut clean, 2, &header_record(0, b"header-0-clean", false));
    push_bytes_field(&mut clean, 2, &header_record(2, b"header-2-clean", false));

    let mut duplicate_root = valid.clone();
    push_varint_field(&mut duplicate_root, 1, 7);
    let mut duplicate_header = valid.clone();
    push_bytes_field(
        &mut duplicate_header,
        2,
        &header_record(2, b"duplicate-index", true),
    );
    let wrong_wire = vec![0x0d, 0x07, 0x00, 0x00, 0x00, 0x00];
    let noncanonical = vec![0x08, 0x87, 0x00];
    let truncated = vec![0x08, 0x07, 0x12, 0x02, 0x08];
    let unterminated_group = vec![0x08, 0x07, 0xdb, 0x05, 0x08, 0x01];
    let mismatched_group = vec![0x08, 0x07, 0xdb, 0x05, 0x08, 0x01, 0xe4, 0x05];
    vec![
        valid,
        clean,
        duplicate_root,
        duplicate_header,
        wrong_wire,
        noncanonical,
        truncated,
        unterminated_group,
        mismatched_group,
    ]
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    let mut value = Vec::new();
    push_varint_field(&mut value, 1, lower);
    push_varint_field(&mut value, 2, upper);
    value
}

fn uid_fixtures() -> Vec<Vec<u8>> {
    let mut valid = Vec::new();
    for (lower, upper) in [(1, 11), (2, 22)] {
        push_bytes_field(&mut valid, 1, &uuid(lower, upper));
    }
    push_varint_field(&mut valid, 2, 0);
    push_varint_field(&mut valid, 2, 1);
    push_varint_field(&mut valid, 3, 0);
    push_varint_field(&mut valid, 3, 1);
    for (lower, upper) in [(101, 111), (102, 122), (103, 133)] {
        push_bytes_field(&mut valid, 4, &uuid(lower, upper));
    }
    for value in [0, 1, 2] {
        push_varint_field(&mut valid, 5, value);
    }
    for value in [0, 1, 2] {
        push_varint_field(&mut valid, 6, value);
    }
    let clean = valid.clone();
    push_bytes_field(&mut valid, 90, b"uid-root-unknown");
    valid.extend_from_slice(UNKNOWN_GROUP);

    let mut duplicate = valid.clone();
    push_varint_field(&mut duplicate, 2, 1);
    let truncated = vec![0x0a, 0x02, 0x08];
    let wrong_wire = vec![0x09, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00];
    let noncanonical = vec![0x08, 0x81, 0x00];
    let mismatched = vec![0x0a, 0x02, 0x08, 0x01, 0x10, 0x00];
    let unterminated_group = vec![0xdb, 0x05, 0x08, 0x01];
    let mismatched_group = vec![0xdb, 0x05, 0x08, 0x01, 0xe4, 0x05];
    vec![
        valid,
        clean,
        duplicate,
        truncated,
        wrong_wire,
        noncanonical,
        mismatched,
        unterminated_group,
        mismatched_group,
    ]
}
