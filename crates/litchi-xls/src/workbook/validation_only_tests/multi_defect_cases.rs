//! Multi-defect worksheet substreams for the frozen worksheet-level
//! first-error matrix.
//!
//! Every case carries two or more defects, or defects hidden behind valid
//! duplicates, at chosen positions, so that the first refusal is the thing
//! under test. The records start after the worksheet `BOF`, as the worksheet
//! walk sees them. This file uses only `std`, so the identical file generated
//! the expected refusals on the base commit (`009d515bef`), whose package walk
//! cannot report them: `parse_workbook` drops a worksheet it cannot parse
//! without surfacing why.

pub(super) type Record = (u16, Vec<u8>);

const EOF: u16 = 0x000A;
const NUMBER: u16 = 0x0203;
const FORMULA: u16 = 0x0006;
const STRING: u16 = 0x0207;
const CONTINUE: u16 = 0x003C;
const ARRAY: u16 = 0x0221;
const SHR_FMLA: u16 = 0x04BC;
const DIMENSIONS: u16 = 0x0200;
const MUL_RK: u16 = 0x00BD;
const MUL_BLANK: u16 = 0x00BE;
const USER_S_VIEW_END: u16 = 0x01AB;
const DVAL: u16 = 0x01B2;
const BAD_XF: u16 = 0x0FFF;
const RESERVED_XF: u16 = 3;

fn head(row: u16, column: u16, xf: u16) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&xf.to_le_bytes());
    payload
}

fn number(row: u16, column: u16, xf: u16) -> Record {
    let mut payload = head(row, column, xf);
    payload.extend_from_slice(&(f64::from(row) + 0.25).to_le_bytes());
    (NUMBER, payload)
}

fn truncated_number(row: u16, column: u16, xf: u16) -> Record {
    let mut payload = head(row, column, xf);
    payload.extend_from_slice(&[0, 0, 0, 0]);
    (NUMBER, payload)
}

fn ptg_exp(anchor: (u16, u16)) -> Vec<u8> {
    let mut tokens = vec![0x01];
    tokens.extend_from_slice(&anchor.0.to_le_bytes());
    tokens.extend_from_slice(&anchor.1.to_le_bytes());
    tokens
}

fn ptg_int() -> Vec<u8> {
    vec![0x1E, 7, 0]
}

/// A `Formula` with a numeric cache (or a pending string cache) and `tokens`.
fn formula(row: u16, column: u16, xf: u16, tokens: &[u8], shared: bool, string: bool) -> Record {
    let mut payload = head(row, column, xf);
    if string {
        payload.extend_from_slice(&[0, 0, 0, 0, 0, 0, 0xff, 0xff]);
    } else {
        payload.extend_from_slice(&1.5_f64.to_le_bytes());
    }
    payload.extend_from_slice(&(if shared { 0x0008_u16 } else { 0 }).to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&u16::try_from(tokens.len()).unwrap_or(0).to_le_bytes());
    payload.extend_from_slice(tokens);
    (FORMULA, payload)
}

/// A compressed `String` record declaring `declared` characters and carrying
/// the bytes of `text`.
fn string(declared: u16, text: &[u8]) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&declared.to_le_bytes());
    payload.push(0);
    payload.extend_from_slice(text);
    (STRING, payload)
}

fn array(first_row: u16, last_row: u16, first_column: u8, last_column: u8) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&first_row.to_le_bytes());
    payload.extend_from_slice(&last_row.to_le_bytes());
    payload.push(first_column);
    payload.push(last_column);
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    let tokens = ptg_int();
    payload.extend_from_slice(&u16::try_from(tokens.len()).unwrap_or(0).to_le_bytes());
    payload.extend_from_slice(&tokens);
    (ARRAY, payload)
}

fn shr_fmla(first_row: u16, last_row: u16, first_column: u8, last_column: u8) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&first_row.to_le_bytes());
    payload.extend_from_slice(&last_row.to_le_bytes());
    payload.push(first_column);
    payload.push(last_column);
    payload.push(0);
    payload.push(1);
    let tokens = ptg_int();
    payload.extend_from_slice(&u16::try_from(tokens.len()).unwrap_or(0).to_le_bytes());
    payload.extend_from_slice(&tokens);
    (SHR_FMLA, payload)
}

fn dimensions(last_row_exclusive: u32, last_column_exclusive: u16) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&last_row_exclusive.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&last_column_exclusive.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    (DIMENSIONS, payload)
}

fn mul_rk(row: u16, first_column: u16, xfs: &[u16], last_column: u16) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&first_column.to_le_bytes());
    for (index, xf) in xfs.iter().enumerate() {
        payload.extend_from_slice(&xf.to_le_bytes());
        let value = u32::try_from(index).unwrap_or(0) + 1;
        payload.extend_from_slice(&((value << 2) | 0x02).to_le_bytes());
    }
    payload.extend_from_slice(&last_column.to_le_bytes());
    (MUL_RK, payload)
}

fn mul_blank(row: u16, first_column: u16, xfs: &[u16], last_column: u16) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&first_column.to_le_bytes());
    for xf in xfs {
        payload.extend_from_slice(&xf.to_le_bytes());
    }
    payload.extend_from_slice(&last_column.to_le_bytes());
    (MUL_BLANK, payload)
}

fn dval(declared: u32) -> Record {
    let mut payload = Vec::new();
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&(-1_i32).to_le_bytes());
    payload.extend_from_slice(&declared.to_le_bytes());
    (DVAL, payload)
}

fn eof() -> Record {
    (EOF, Vec::new())
}

/// The cases, each a name and the worksheet records after `BOF`. `ok` is a
/// valid cell XF index of the workbook whose formatting table parses them.
pub(super) fn cases(ok: u16) -> Vec<(&'static str, Vec<Record>)> {
    let a = (4_u16, 0_u16);
    vec![
        (
            "dup-then-bad-xf",
            vec![
                number(1, 1, ok),
                number(1, 1, ok),
                number(2, 2, BAD_XF),
                eof(),
            ],
        ),
        (
            "bad-xf-then-truncated",
            vec![number(1, 1, BAD_XF), truncated_number(2, 2, ok), eof()],
        ),
        (
            "truncated-then-bad-xf",
            vec![truncated_number(1, 1, ok), number(2, 2, BAD_XF), eof()],
        ),
        (
            "reserved-xf-then-bad-xf",
            vec![number(1, 1, RESERVED_XF), number(2, 2, BAD_XF), eof()],
        ),
        (
            "pending-string-formula-then-stray-string",
            vec![
                formula(1, 1, ok, &ptg_int(), false, true),
                number(1, 2, ok),
                string(1, b"x"),
                eof(),
            ],
        ),
        (
            "stray-string-then-orphan",
            vec![
                string(1, b"x"),
                formula(5, 5, ok, &ptg_exp((5, 5)), false, false),
                eof(),
            ],
        ),
        (
            "string-continuation-not-continue-then-bad-xf",
            vec![
                formula(1, 1, ok, &ptg_int(), false, true),
                string(5, b"ab"),
                number(2, 2, BAD_XF),
                (CONTINUE, vec![0, b'c']),
                eof(),
            ],
        ),
        (
            "orphan-then-array-without-dimensions",
            vec![
                formula(2, 2, ok, &ptg_exp((2, 2)), false, false),
                formula(a.0, a.1, ok, &ptg_exp(a), false, false),
                array(4, 5, 0, 0),
                formula(5, 0, ok, &ptg_exp(a), false, false),
                eof(),
            ],
        ),
        (
            "array-without-dimensions-then-duplicate-member",
            vec![
                formula(a.0, a.1, ok, &ptg_exp(a), false, false),
                array(4, 5, 0, 0),
                formula(5, 0, ok, &ptg_exp(a), false, false),
                number(5, 0, ok),
                eof(),
            ],
        ),
        (
            "array-duplicate-anchor-then-missing-member",
            vec![
                dimensions(10, 5),
                number(4, 0, ok),
                formula(a.0, a.1, ok, &ptg_exp(a), false, false),
                array(4, 6, 0, 0),
                formula(5, 0, ok, &ptg_exp(a), false, false),
                eof(),
            ],
        ),
        (
            "array-duplicate-member-before-missing-member",
            vec![
                dimensions(10, 5),
                formula(a.0, a.1, ok, &ptg_exp(a), false, false),
                array(4, 6, 0, 0),
                formula(5, 0, ok, &ptg_exp(a), false, false),
                number(5, 0, ok),
                eof(),
            ],
        ),
        (
            "second-array-duplicate-member",
            vec![
                dimensions(10, 5),
                formula(1, 1, ok, &ptg_exp((1, 1)), false, false),
                array(1, 2, 1, 1),
                formula(2, 1, ok, &ptg_exp((1, 1)), false, false),
                formula(4, 1, ok, &ptg_exp((4, 1)), false, false),
                array(4, 5, 1, 1),
                formula(5, 1, ok, &ptg_exp((4, 1)), false, false),
                number(5, 1, ok),
                eof(),
            ],
        ),
        (
            "array-outside-dimensions-with-duplicates",
            vec![
                dimensions(3, 2),
                number(0, 0, ok),
                number(0, 0, ok),
                formula(1, 1, ok, &ptg_exp((1, 1)), false, false),
                array(1, 4, 1, 1),
                formula(2, 1, ok, &ptg_exp((1, 1)), false, false),
                formula(3, 1, ok, &ptg_exp((1, 1)), false, false),
                formula(4, 1, ok, &ptg_exp((1, 1)), false, false),
                eof(),
            ],
        ),
        (
            "shrfmla-without-formula-then-bad-xf",
            vec![
                number(1, 1, ok),
                shr_fmla(1, 2, 1, 1),
                number(2, 2, BAD_XF),
                eof(),
            ],
        ),
        (
            "duplicate-anchor-claims-second-companion",
            vec![
                formula(1, 1, ok, &ptg_exp((1, 1)), true, false),
                shr_fmla(1, 2, 1, 1),
                formula(2, 1, ok, &ptg_exp((1, 1)), true, false),
                formula(1, 1, ok, &ptg_exp((1, 1)), false, false),
                array(1, 1, 1, 1),
                eof(),
            ],
        ),
        (
            "mulblank-bad-xf-then-mulrk-bad-range",
            vec![
                mul_blank(3, 1, &[ok, BAD_XF, ok], 3),
                mul_rk(4, 2, &[ok, ok], 9),
                eof(),
            ],
        ),
        (
            "mulrk-bad-range-then-bad-xf",
            vec![mul_rk(4, 2, &[ok, ok], 9), number(5, 5, BAD_XF), eof()],
        ),
        (
            "mulrk-over-number-then-reserved-xf",
            vec![
                number(4, 2, ok),
                mul_rk(4, 1, &[ok, ok], 2),
                mul_blank(4, 2, &[ok, RESERVED_XF], 3),
                eof(),
            ],
        ),
        (
            "unmaterialized-member-behind-orphan",
            vec![
                dimensions(10, 10),
                formula(1, 0, ok, &ptg_exp((1, 0)), false, false),
                array(1, 2, 0, 0),
                number(7, 7, ok),
                formula(3, 3, ok, &ptg_exp((9, 9)), false, false),
                formula(2, 0, ok, &ptg_exp((1, 0)), false, true),
            ],
        ),
        (
            "unmaterialized-member",
            vec![
                dimensions(10, 10),
                formula(1, 0, ok, &ptg_exp((1, 0)), false, false),
                array(1, 2, 0, 0),
                number(7, 7, ok),
                formula(2, 0, ok, &ptg_exp((1, 0)), false, true),
            ],
        ),
        (
            "non-formula-member-with-duplicates",
            vec![
                dimensions(10, 10),
                number(6, 6, ok),
                number(6, 6, ok),
                number(2, 0, ok),
                formula(1, 0, ok, &ptg_exp((1, 0)), false, false),
                array(1, 2, 0, 0),
                formula(2, 0, ok, &ptg_exp((1, 0)), false, true),
            ],
        ),
        (
            "outside-grid-duplicates-then-bad-xf",
            vec![
                number(0, 300, ok),
                number(0, 300, ok),
                number(0, 301, BAD_XF),
                eof(),
            ],
        ),
        (
            "custom-view-end-without-begin-then-bad-xf",
            vec![
                number(1, 1, ok),
                (USER_S_VIEW_END, vec![0, 0]),
                number(2, 2, BAD_XF),
                eof(),
            ],
        ),
        (
            "dval-cut-short-by-a-cell-then-bad-xf",
            vec![dval(2), number(1, 1, ok), number(2, 2, BAD_XF), eof()],
        ),
        (
            "dval-cut-short-by-eof-after-duplicates",
            vec![number(1, 1, ok), number(1, 1, ok), dval(1), eof()],
        ),
        (
            "valid-duplicates-shared-and-array",
            vec![
                dimensions(10, 5),
                number(0, 0, ok),
                number(0, 0, ok),
                formula(1, 1, ok, &ptg_exp((1, 1)), true, false),
                shr_fmla(1, 2, 1, 1),
                formula(2, 1, ok, &ptg_exp((1, 1)), true, false),
                formula(a.0, a.1, ok, &ptg_exp(a), false, false),
                array(4, 5, 0, 0),
                formula(5, 0, ok, &ptg_exp(a), false, false),
                number(9, 300, ok),
                eof(),
            ],
        ),
        (
            "no-eof-pending-string-formula-after-duplicates",
            vec![
                number(1, 1, ok),
                number(1, 1, ok),
                formula(3, 3, ok, &ptg_int(), false, true),
            ],
        ),
    ]
}
