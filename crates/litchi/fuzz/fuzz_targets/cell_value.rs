#![no_main]

//! Bounded, allocation-free semantic fuzzing for the shared Numbers cell
//! value reader.
//!
//! Arbitrary bytes exercise the raw parser directly. Small stack-backed
//! profiles keep the semantic routes reachable even when the input is not a
//! native cell: every supported scalar shape is covered for BNC and selected
//! pre-BNC versions, with formula precedence, zero identifiers, comments,
//! truncation, and an unsupported profile.

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_numbers_wire::cell_value::decode_cell_value;

const MAX_INPUT_BYTES: usize = 1_024;
const FIXTURE_BYTES: usize = 128;

const BNC_VERSION: u8 = 5;
const BNC_TYPE_EMPTY: u8 = 0;
const BNC_TYPE_NUMBER: u8 = 2;
const BNC_TYPE_TEXT: u8 = 3;
const BNC_TYPE_DATE: u8 = 5;
const BNC_TYPE_BOOLEAN: u8 = 6;
const BNC_TYPE_DURATION: u8 = 7;
const BNC_TYPE_ERROR: u8 = 8;
const BNC_TYPE_RICH_TEXT_OR_NUMBER: u8 = 9;

const BNC_FLAG_NUMBER: u32 = 0x0000_0002;
const BNC_FLAG_DATE: u32 = 0x0000_0004;
const BNC_FLAG_STRING: u32 = 0x0000_0008;
const BNC_FLAG_RICH_TEXT: u32 = 0x0000_0010;
const BNC_FLAG_FORMULA: u32 = 0x0000_0200;
const BNC_FLAG_FORMULA_ERROR: u32 = 0x0000_0800;
const BNC_FLAG_COMMENT: u32 = 0x0008_0000;

const PRE_FLAG_FORMULA: u32 = 0x0000_0008;
const PRE_FLAG_FORMULA_ERROR: u32 = 0x0000_0100;
const PRE_FLAG_RICH_TEXT: u32 = 0x0000_0200;
const PRE_FLAG_COMMENT: u32 = 0x0000_1000;
const PRE_FLAG_STRING: u32 = 0x0000_0010;
const PRE_FLAG_NUMBER: u32 = 0x0000_0020;
const PRE_FLAG_DATE: u32 = 0x0000_0040;

/// A tiny fixed-capacity byte builder. The decoder under test receives only
/// this borrowed slice; no owned AST or package model is constructed.
#[derive(Clone, Copy)]
struct SmallBuffer {
    bytes: [u8; FIXTURE_BYTES],
    len: usize,
}

impl SmallBuffer {
    fn new() -> Self {
        Self {
            bytes: [0; FIXTURE_BYTES],
            len: 0,
        }
    }

    fn push(&mut self, byte: u8) {
        assert!(self.len < self.bytes.len(), "cell fixture exceeded its cap");
        self.bytes[self.len] = byte;
        self.len += 1;
    }

    fn extend(&mut self, bytes: &[u8]) {
        let end = self
            .len
            .checked_add(bytes.len())
            .expect("cell fixture length overflow");
        assert!(end <= self.bytes.len(), "cell fixture exceeded its cap");
        self.bytes[self.len..end].copy_from_slice(bytes);
        self.len = end;
    }

    fn as_slice(&self) -> &[u8] {
        &self.bytes[..self.len]
    }
}

fuzz_target!(|data: &[u8]| {
    let bounded = &data[..data.len().min(MAX_INPUT_BYTES)];
    exercise_source(bounded);

    // Keep all semantic controls small and deterministic while still letting
    // each input byte select a different native value shape and identifier.
    for case in 0..=21_u8 {
        let fixture = semantic_fixture(data, case);
        exercise_source(fixture.as_slice());

        // A one-byte truncation drives the same bounded profile through the
        // typed malformed-input path without feeding an unbounded slice.
        if fixture.len > 1 {
            exercise_source(&fixture.as_slice()[..fixture.len - 1]);
        }
    }
});

fn exercise_source(source: &[u8]) {
    match decode_cell_value(source) {
        Ok(value) => {
            black_box((value.value, value.comment_identifier));
        },
        Err(error) => {
            black_box(error);
        },
    }
}

fn semantic_fixture(data: &[u8], case: u8) -> SmallBuffer {
    let mut fixture = SmallBuffer::new();
    let scalar = finite_scalar(data, 1);
    let identifier = control_u32(data, 5);
    let comment_identifier = control_u32(data, 9);
    let include_comment = control_byte(data, 13) & 1 != 0;

    match case {
        0 => bnc(&mut fixture, BNC_TYPE_EMPTY, 0, scalar, identifier, None),
        1 => bnc(
            &mut fixture,
            BNC_TYPE_NUMBER,
            BNC_FLAG_NUMBER,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        2 => bnc(
            &mut fixture,
            BNC_TYPE_DATE,
            BNC_FLAG_DATE,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        3 => bnc(
            &mut fixture,
            BNC_TYPE_BOOLEAN,
            BNC_FLAG_NUMBER,
            if control_byte(data, 2) & 1 == 0 {
                0.0
            } else {
                1.0
            },
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        4 => bnc(
            &mut fixture,
            BNC_TYPE_DURATION,
            BNC_FLAG_NUMBER,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        5 => bnc(
            &mut fixture,
            BNC_TYPE_TEXT,
            BNC_FLAG_STRING,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        6 => bnc(
            &mut fixture,
            BNC_TYPE_RICH_TEXT_OR_NUMBER,
            BNC_FLAG_RICH_TEXT,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        7 => bnc(
            &mut fixture,
            BNC_TYPE_NUMBER,
            BNC_FLAG_FORMULA | BNC_FLAG_STRING,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        8 => bnc(
            &mut fixture,
            BNC_TYPE_ERROR,
            BNC_FLAG_FORMULA_ERROR,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        9 => bnc(
            &mut fixture,
            BNC_TYPE_EMPTY,
            BNC_FLAG_COMMENT,
            scalar,
            identifier,
            Some(comment_identifier),
        ),
        10 => bnc(
            &mut fixture,
            0x7f,
            0,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        11 => bnc(
            &mut fixture,
            BNC_TYPE_NUMBER,
            0x8000_0000,
            scalar,
            identifier,
            None,
        ),
        12 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            0,
            0,
            scalar,
            identifier,
            None,
        ),
        13 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            2,
            PRE_FLAG_NUMBER,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        14 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            5,
            PRE_FLAG_DATE,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        15 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            6,
            PRE_FLAG_NUMBER,
            if control_byte(data, 2) & 1 == 0 {
                0.0
            } else {
                1.0
            },
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        16 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            7,
            PRE_FLAG_NUMBER,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        17 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            3,
            PRE_FLAG_STRING,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        18 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            9,
            PRE_FLAG_RICH_TEXT,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        19 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            2,
            PRE_FLAG_FORMULA | PRE_FLAG_STRING,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        20 => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            8,
            PRE_FLAG_FORMULA_ERROR,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
        _ => pre_bnc(
            &mut fixture,
            control_byte(data, 0) % 5,
            0x7f,
            0,
            scalar,
            identifier,
            include_comment.then_some(comment_identifier),
        ),
    }

    append_tail(&mut fixture, data);
    fixture
}

fn bnc(
    output: &mut SmallBuffer,
    cell_type: u8,
    flags: u32,
    scalar: f64,
    identifier: u32,
    comment_identifier: Option<u32>,
) {
    output.extend(&[BNC_VERSION, cell_type, 0, 0, 0, 0, 0, 0]);
    output.extend(&flags.to_le_bytes());

    // The field order mirrors the fixed BNC layout. These profiles use only
    // the scalar and semantic-reference fields, so no generated protobuf AST
    // is involved.
    if flags & BNC_FLAG_NUMBER != 0 {
        output.extend(&scalar.to_le_bytes());
    }
    if flags & BNC_FLAG_DATE != 0 {
        output.extend(&scalar.to_le_bytes());
    }
    if flags & BNC_FLAG_STRING != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & BNC_FLAG_RICH_TEXT != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & BNC_FLAG_FORMULA != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & BNC_FLAG_FORMULA_ERROR != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & BNC_FLAG_COMMENT != 0 {
        output.extend(&comment_identifier.unwrap_or(identifier).to_le_bytes());
    }
}

fn pre_bnc(
    output: &mut SmallBuffer,
    version: u8,
    cell_type: u8,
    flags: u32,
    scalar: f64,
    identifier: u32,
    comment_identifier: Option<u32>,
) {
    let header_len = if version <= 1 { 8 } else { 12 };
    output.push(version);
    if version == 4 {
        output.push(cell_type);
        output.extend(&[0xa5, 0x5a]);
    } else {
        output.extend(&[0x5a, cell_type, 0xa5]);
    }
    if version <= 1 {
        output.extend(&(flags as u16).to_le_bytes());
        output.extend(&[0xc3, 0x3d]);
    } else {
        output.extend(&flags.to_le_bytes());
        output.extend(&[0x19, 0x91, 0x71, 0x17]);
    }
    debug_assert_eq!(output.len, header_len);

    // The legacy reader consumes these slots in order, including fields that
    // are not projected by the shared semantic value.
    if flags & PRE_FLAG_FORMULA != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & PRE_FLAG_FORMULA_ERROR != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & PRE_FLAG_RICH_TEXT != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & PRE_FLAG_COMMENT != 0 {
        output.extend(&comment_identifier.unwrap_or(identifier).to_le_bytes());
    }
    if flags & PRE_FLAG_STRING != 0 {
        output.extend(&identifier.to_le_bytes());
    }
    if flags & PRE_FLAG_NUMBER != 0 {
        output.extend(&scalar.to_le_bytes());
    }
    if flags & PRE_FLAG_DATE != 0 {
        output.extend(&scalar.to_le_bytes());
    }
}

fn append_tail(output: &mut SmallBuffer, data: &[u8]) {
    let count = (usize::from(control_byte(data, 14)) % 5).min(data.len().saturating_sub(15));
    if count == 0 {
        output.extend(&[0xca, 0xfe]);
        return;
    }
    output.extend(&data[15..15 + count]);
}

fn finite_scalar(data: &[u8], offset: usize) -> f64 {
    let raw = u16::from_le_bytes([control_byte(data, offset), control_byte(data, offset + 1)]);
    f64::from(raw % 4096) / 16.0
}

fn control_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control_byte(data, offset),
        control_byte(data, offset + 1),
        control_byte(data, offset + 2),
        control_byte(data, offset + 3),
    ])
}

fn control_byte(data: &[u8], offset: usize) -> u8 {
    data.get(offset)
        .copied()
        .unwrap_or((offset as u8).wrapping_mul(37).wrapping_add(11))
}
