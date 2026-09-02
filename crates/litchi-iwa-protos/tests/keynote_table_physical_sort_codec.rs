use litchi_iwa_protos::table_physical_sort_codec::{
    DecodeOptions, HeaderRowMove, RowMove, RowUidPermutation, decode_column_row_uid_map,
    plan_header_storage_bucket_rows, plan_tile_rows_rewrite, rewrite_column_row_uid_map,
    rewrite_header_storage_bucket_rows, rewrite_tile_rows,
};

fn varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
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
    varint(output, u64::from(field) << 3);
    varint(output, value);
}

fn field_bytes(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
    varint(output, (u64::from(field) << 3) | 2);
    varint(
        output,
        u64::try_from(payload.len()).expect("test payload length fits"),
    );
    output.extend_from_slice(payload);
}

fn row(index: u32, marker: u8) -> Vec<u8> {
    row_with_unknown(index, marker, true)
}

fn row_with_unknown(index: u32, marker: u8, include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, u64::from(index));
    field_varint(&mut output, 2, 1);
    field_bytes(&mut output, 3, &[marker]);
    field_bytes(&mut output, 4, &[marker, 0x42]);
    if include_unknown {
        // Unknown nested bytes remain source-preserved for a no-op read.
        field_bytes(&mut output, 99, &[0x7a, marker]);
    }
    output
}

fn row_with_payload_len_clean(index: u32, payload_len: usize) -> Vec<u8> {
    row_with_payload_len_with_unknown(index, payload_len, false)
}

fn row_with_payload_len_with_unknown(
    index: u32,
    payload_len: usize,
    include_unknown: bool,
) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, u64::from(index));
    field_varint(&mut output, 2, 1);
    field_bytes(&mut output, 3, &[0xa0]);
    field_bytes(&mut output, 4, &[0xb0]);
    if include_unknown {
        let fixed_len = output.len() + 3;
        assert!(payload_len >= fixed_len);
        field_bytes(&mut output, 99, &vec![0x5a; payload_len - fixed_len]);
    } else {
        // Field 6 is a known opaque storage buffer.  Its one-byte key and
        // one-byte length prefix make the fixture exactly payload_len bytes.
        let fixed_len = output.len() + 2;
        assert!(payload_len >= fixed_len);
        field_bytes(&mut output, 6, &vec![0x5a; payload_len - fixed_len]);
    }
    assert_eq!(output.len(), payload_len);
    output
}

fn tile() -> Vec<u8> {
    tile_with_unknowns(true)
}

fn tile_clean() -> Vec<u8> {
    tile_with_unknowns(false)
}

fn tile_with_unknowns(include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, 3);
    field_varint(&mut output, 2, 2);
    field_varint(&mut output, 3, 2);
    field_varint(&mut output, 4, 3);
    if include_unknown {
        field_varint(&mut output, 90, 7);
    }
    field_bytes(&mut output, 5, &row_with_unknown(0, b'a', include_unknown));
    if include_unknown {
        field_varint(&mut output, 91, 8);
    }
    field_bytes(&mut output, 5, &row_with_unknown(2, b'b', include_unknown));
    output
}

fn header(index: u32, marker: u8) -> Vec<u8> {
    header_with_unknown(index, marker, true)
}

fn header_with_unknown(index: u32, marker: u8, include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, u64::from(index));
    // fixed32 size_bits
    varint(&mut output, (2 << 3) | 5);
    output.extend_from_slice(&u32::from(marker).to_le_bytes());
    field_varint(&mut output, 3, 0);
    field_varint(&mut output, 4, 0);
    if include_unknown {
        field_bytes(&mut output, 77, &[marker]);
    }
    output
}

fn header_bucket() -> Vec<u8> {
    header_bucket_with_unknowns(true)
}

fn header_bucket_clean() -> Vec<u8> {
    header_bucket_with_unknowns(false)
}

fn header_bucket_with_unknowns(include_unknown: bool) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, 4);
    if include_unknown {
        field_varint(&mut output, 93, 9);
    }
    field_bytes(
        &mut output,
        2,
        &header_with_unknown(0, b'x', include_unknown),
    );
    field_bytes(
        &mut output,
        2,
        &header_with_unknown(10, b'y', include_unknown),
    );
    output
}

fn uuid(lower: u64, upper: u64) -> Vec<u8> {
    let mut output = Vec::new();
    field_varint(&mut output, 1, lower);
    field_varint(&mut output, 2, upper);
    output
}

fn uid_map() -> Vec<u8> {
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
    output
}

fn options() -> DecodeOptions {
    DecodeOptions::new(1 << 20, 10_000, 4 << 20, 32, 1_000, 1_000, 1 << 20, 1 << 20)
}

#[test]
fn tile_rows_rewrite_is_sparse_and_rejects_unknown_mutation() {
    let unknown_source = tile();
    let (identity, identity_report) =
        rewrite_tile_rows(&unknown_source, 256, &[], options()).unwrap();
    assert_eq!(identity, unknown_source);
    assert_eq!(identity_report.changed_records(), 0);
    let unknown_moves = [RowMove::new(0, 2), RowMove::new(2, 0)];
    assert!(rewrite_tile_rows(&unknown_source, 256, &unknown_moves, options()).is_err());

    let source = tile_clean();
    let moves = [RowMove::new(0, 2), RowMove::new(2, 0)];
    let (rewritten, report) = rewrite_tile_rows(&source, 256, &moves, options()).unwrap();
    assert_eq!(report.changed_records(), 2);
    assert!(
        rewritten
            .windows(3)
            .any(|window| window == [0x1a, 0x01, b'b'])
    );
    let plan = plan_tile_rows_rewrite(&rewritten, 256, &[], options()).unwrap();
    let indices = plan
        .row_records()
        .map(|record| record.snapshot().tile_row_index())
        .collect::<Vec<_>>();
    assert_eq!(indices, [0, 2]);
}

#[test]
fn header_rows_allow_sparse_records_and_reject_unknown_mutation() {
    let unknown_source = header_bucket();
    let (identity, identity_report) =
        rewrite_header_storage_bucket_rows(&unknown_source, 32, &[], options()).unwrap();
    assert_eq!(identity, unknown_source);
    assert_eq!(identity_report.changed_records(), 0);
    let unknown_moves = [HeaderRowMove::new(0, 10), HeaderRowMove::new(10, 0)];
    assert!(
        rewrite_header_storage_bucket_rows(&unknown_source, 32, &unknown_moves, options(),)
            .is_err()
    );

    let source = header_bucket_clean();
    let moves = [HeaderRowMove::new(0, 10), HeaderRowMove::new(10, 0)];
    let (rewritten, report) =
        rewrite_header_storage_bucket_rows(&source, 32, &moves, options()).unwrap();
    assert_eq!(report.changed_records(), 2);
    let plan = plan_header_storage_bucket_rows(&rewritten, 32, &[], options()).unwrap();
    let indices = plan
        .records()
        .map(|record| record.snapshot().index())
        .collect::<Vec<_>>();
    assert_eq!(indices, [0, 10]);
}

#[test]
fn row_uid_rewrite_updates_both_inverse_arrays() {
    let source = uid_map();
    let (snapshot, report) = decode_column_row_uid_map(&source, 2, 3, options()).unwrap();
    assert_eq!(snapshot.row_uid_for_index(), [0, 1, 2]);
    assert_eq!(snapshot.row_index_for_uid(), [0, 1, 2]);
    assert!(report.elements() >= 15);

    let permutation = RowUidPermutation::new(&[2, 0, 1]).unwrap();
    let (rewritten, rewrite_report) =
        rewrite_column_row_uid_map(&source, 2, 3, &permutation, options()).unwrap();
    assert!(rewrite_report.changed_records() > 0);
    let (snapshot, _) = decode_column_row_uid_map(&rewritten, 2, 3, options()).unwrap();
    assert_eq!(snapshot.row_uid_for_index(), [1, 2, 0]);
    assert_eq!(snapshot.row_index_for_uid(), [2, 0, 1]);
    let (identity, _) = rewrite_column_row_uid_map(
        &rewritten,
        2,
        3,
        &RowUidPermutation::identity(3).unwrap(),
        options(),
    )
    .unwrap();
    assert_eq!(identity, rewritten);
}

#[test]
fn malformed_wire_and_duplicate_admission_fails_closed() {
    let mut duplicate_tile = tile();
    let duplicate = row(0, b'z');
    field_bytes(&mut duplicate_tile, 5, &duplicate);
    assert!(plan_tile_rows_rewrite(&duplicate_tile, 256, &[], options()).is_err());

    let mut duplicate_header = header_bucket();
    field_bytes(&mut duplicate_header, 2, &header(0, b'z'));
    assert!(plan_header_storage_bucket_rows(&duplicate_header, 32, &[], options()).is_err());

    let mut wrong_wire = tile();
    wrong_wire[0] = 0x09;
    assert!(plan_tile_rows_rewrite(&wrong_wire, 256, &[], options()).is_err());

    let mut overlong = tile();
    // Replace the first scalar's canonical `0x18` payload with an overlong
    // varint. The strict key/value scanner rejects it before parity.
    overlong.splice(0..2, [0x08, 0x80, 0x00]);
    assert!(plan_tile_rows_rewrite(&overlong, 256, &[], options()).is_err());

    let mut duplicate_uid = uid_map();
    field_varint(&mut duplicate_uid, 6, 0);
    assert!(decode_column_row_uid_map(&duplicate_uid, 2, 3, options()).is_err());
}

#[test]
fn limits_are_typed_before_staging() {
    let constrained = DecodeOptions::new(8, 10, 10, 8, 1, 1, 8, 8);
    let error = plan_tile_rows_rewrite(&tile(), 256, &[], constrained).unwrap_err();
    assert!(error.resource_limit().is_some());
}

#[test]
fn nested_length_prefix_changes_are_preflighted_exactly() {
    let mut source = Vec::new();
    field_varint(&mut source, 1, 3);
    field_varint(&mut source, 2, 2);
    field_varint(&mut source, 3, 2);
    field_varint(&mut source, 4, 129);
    field_bytes(&mut source, 5, &row_with_payload_len_clean(0, 127));
    field_bytes(&mut source, 5, &row_with_payload_len_clean(128, 127));

    let moves = [RowMove::new(0, 129), RowMove::new(128, 0)];
    let (rewritten, report) = rewrite_tile_rows(&source, 256, &moves, options()).unwrap();
    assert_eq!(rewritten.len(), report.output_bytes());
    assert_eq!(rewritten.len(), source.len() + 1);
    assert_eq!(
        plan_tile_rows_rewrite(&rewritten, 256, &[], options())
            .unwrap()
            .tile()
            .num_rows(),
        130
    );
}
