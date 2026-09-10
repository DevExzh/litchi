//! Independent semantic coverage for the shared Numbers cell-value reader.
//!
//! The fixtures below are handwritten wire payloads.  They deliberately do
//! not use either production cell encoder, so the tests continue to describe
//! the storage contract while the package readers migrate to
//! `cell_value::decode_cell_value`.

use std::mem::{needs_drop, size_of};

use litchi_iwa_common::formula::FiniteF64;
use litchi_numbers_wire::cell_value::{
    CellValueSource, DecodeError, ValueSource, decode_cell_value,
};

const PRE_FORMULA: u32 = 0x0000_0008;
const PRE_STRING: u32 = 0x0000_0010;
const PRE_NUMBER: u32 = 0x0000_0020;
const PRE_DATE: u32 = 0x0000_0040;
const PRE_FORMULA_ERROR: u32 = 0x0000_0100;
const PRE_RICH_TEXT: u32 = 0x0000_0200;
const PRE_COMMENT: u32 = 0x0000_1000;

const PRE_FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0002, 4),
    (0x0000_0080, 4),
    (0x0000_0400, 4),
    (0x0000_0800, 4),
    (0x0000_0004, 4),
    (PRE_FORMULA, 4),
    (PRE_FORMULA_ERROR, 4),
    (PRE_RICH_TEXT, 4),
    (PRE_COMMENT, 4),
    (0x0000_2000, 4),
    (PRE_STRING, 4),
    (PRE_NUMBER, 8),
    (PRE_DATE, 8),
    (0x0001_0000, 4),
    (0x0008_0000, 4),
    (0x0002_0000, 4),
    (0x0004_0000, 4),
    (0x0010_0000, 4),
    (0x0020_0000, 4),
    (0x0040_0000, 4),
    (0x0080_0000, 4),
];

const BNC_DECIMAL: u32 = 0x0000_0001;
const BNC_NUMBER: u32 = 0x0000_0002;
const BNC_DATE: u32 = 0x0000_0004;
const BNC_STRING: u32 = 0x0000_0008;
const BNC_RICH_TEXT: u32 = 0x0000_0010;
const BNC_FORMULA: u32 = 0x0000_0200;
const BNC_FORMULA_ERROR: u32 = 0x0000_0800;
const BNC_COMMENT: u32 = 0x0008_0000;

const BNC_FIELD_LAYOUT: &[(u32, usize)] = &[
    (BNC_DECIMAL, 16),
    (BNC_NUMBER, 8),
    (BNC_DATE, 8),
    (BNC_STRING, 4),
    (BNC_RICH_TEXT, 4),
    (0x0000_0020, 4),
    (0x0000_0040, 4),
    (0x0000_0080, 4),
    (0x0000_0100, 4),
    (BNC_FORMULA, 4),
    (0x0000_0400, 4),
    (BNC_FORMULA_ERROR, 4),
    (0x0000_1000, 4),
    (0x0000_2000, 4),
    (0x0000_4000, 4),
    (0x0000_8000, 4),
    (0x0001_0000, 4),
    (0x0002_0000, 4),
    (0x0004_0000, 4),
    (BNC_COMMENT, 4),
    (0x0010_0000, 4),
];

fn finite(value: f64) -> FiniteF64 {
    FiniteF64::new(value).expect("test scalar must be finite")
}

fn assert_value(actual: ValueSource, expected: ValueSource) {
    match (actual, expected) {
        (ValueSource::Empty, ValueSource::Empty)
        | (ValueSource::Boolean(false), ValueSource::Boolean(false))
        | (ValueSource::Boolean(true), ValueSource::Boolean(true)) => {},
        (ValueSource::Number(actual), ValueSource::Number(expected))
        | (ValueSource::Date(actual), ValueSource::Date(expected))
        | (ValueSource::Duration(actual), ValueSource::Duration(expected)) => {
            assert_eq!(actual.get().to_bits(), expected.get().to_bits());
        },
        (ValueSource::Text(actual), ValueSource::Text(expected))
        | (ValueSource::RichText(actual), ValueSource::RichText(expected))
        | (ValueSource::Formula(actual), ValueSource::Formula(expected)) => {
            assert_eq!(actual, expected);
        },
        (ValueSource::Error(actual), ValueSource::Error(expected)) => {
            assert_eq!(actual, expected);
        },
        (actual, expected) => panic!("value mismatch: got {actual:?}, expected {expected:?}"),
    }
}

fn assert_decoded(source: &[u8], expected: ValueSource, comment_identifier: Option<u32>) {
    let decoded = decode_cell_value(source).expect("handwritten cell should decode");
    assert_value(decoded.value, expected);
    assert_eq!(decoded.comment_identifier, comment_identifier);
}

fn pre_header(version: u8, cell_type: u8, flags: u32) -> Vec<u8> {
    let header_length = usize::from(version > 1) * 4 + 8;
    let mut bytes = vec![0; header_length];
    bytes[0] = version;
    if version == 4 {
        bytes[1] = cell_type;
        bytes[2..4].copy_from_slice(&[0xa5, 0x5a]);
    } else {
        bytes[1] = 0x5a;
        bytes[2] = cell_type;
        bytes[3] = 0xa5;
    }
    if version <= 1 {
        bytes[4..6].copy_from_slice(&(flags as u16).to_le_bytes());
        bytes[6..8].copy_from_slice(&[0xc3, 0x3d]);
    } else {
        bytes[4..8].copy_from_slice(&flags.to_le_bytes());
        bytes[8..12].copy_from_slice(&[0x19, 0x91, 0x71, 0x17]);
    }
    bytes
}

#[allow(clippy::too_many_arguments)]
fn pre_cell(
    version: u8,
    cell_type: u8,
    flags: u32,
    number: Option<f64>,
    date: Option<f64>,
    string_identifier: Option<u32>,
    rich_text_identifier: Option<u32>,
    formula_identifier: Option<u32>,
    formula_error_identifier: Option<u32>,
    comment_identifier: Option<u32>,
) -> Vec<u8> {
    let mut bytes = pre_header(version, cell_type, flags);
    for &(flag, length) in PRE_FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        let field = match flag {
            PRE_FORMULA => formula_identifier
                .unwrap_or(0x1020_3040)
                .to_le_bytes()
                .to_vec(),
            PRE_FORMULA_ERROR => formula_error_identifier
                .unwrap_or(0x5060_7080)
                .to_le_bytes()
                .to_vec(),
            PRE_RICH_TEXT => rich_text_identifier
                .unwrap_or(0x1122_3344)
                .to_le_bytes()
                .to_vec(),
            PRE_COMMENT => comment_identifier
                .unwrap_or(0x99aa_bbcc)
                .to_le_bytes()
                .to_vec(),
            PRE_STRING => string_identifier
                .unwrap_or(0x5566_7788)
                .to_le_bytes()
                .to_vec(),
            PRE_NUMBER => number.unwrap_or(12.5).to_le_bytes().to_vec(),
            PRE_DATE => date.unwrap_or(345.25).to_le_bytes().to_vec(),
            _ => {
                let value = flag.wrapping_mul(3).wrapping_add(7).to_le_bytes();
                value[..length].to_vec()
            },
        };
        assert_eq!(field.len(), length, "legacy field layout is inconsistent");
        bytes.extend_from_slice(&field);
    }
    bytes
}

// This is the finite subset of decimal128 needed by the independent corpus.
// The coefficient and base-10 exponent are encoded directly instead of using
// the production decimal encoder.
fn decimal128(coefficient: u128, exponent: i32, negative: bool) -> [u8; 16] {
    let biased_exponent = u128::try_from(0x1820_i32 + exponent).expect("test exponent is valid");
    let mut encoded = coefficient | (biased_exponent << 113);
    if negative {
        encoded |= 1_u128 << 127;
    }
    encoded.to_le_bytes()
}

fn nonfinite_decimal128() -> [u8; 16] {
    let mut bytes = [0; 16];
    bytes[14] = 0xff;
    bytes[15] = 0x7f;
    bytes
}

#[allow(clippy::too_many_arguments)]
fn bnc_cell(
    cell_type: u8,
    flags: u32,
    decimal: Option<[u8; 16]>,
    number: Option<f64>,
    date: Option<f64>,
    string_identifier: Option<u32>,
    rich_text_identifier: Option<u32>,
    formula_identifier: Option<u32>,
    formula_error_identifier: Option<u32>,
    comment_identifier: Option<u32>,
) -> Vec<u8> {
    let mut bytes = vec![5, cell_type, 0xa5, 0x5a, 0, 0, 0, 0];
    bytes.extend_from_slice(&flags.to_le_bytes());
    for &(flag, length) in BNC_FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        let field = match flag {
            BNC_DECIMAL => decimal
                .unwrap_or_else(|| decimal128(125, -1, false))
                .to_vec(),
            BNC_NUMBER => number.unwrap_or(12.5).to_le_bytes().to_vec(),
            BNC_DATE => date.unwrap_or(345.25).to_le_bytes().to_vec(),
            BNC_STRING => string_identifier
                .unwrap_or(0x5566_7788)
                .to_le_bytes()
                .to_vec(),
            BNC_RICH_TEXT => rich_text_identifier
                .unwrap_or(0x1122_3344)
                .to_le_bytes()
                .to_vec(),
            BNC_FORMULA => formula_identifier
                .unwrap_or(0x1020_3040)
                .to_le_bytes()
                .to_vec(),
            BNC_FORMULA_ERROR => formula_error_identifier
                .unwrap_or(0x5060_7080)
                .to_le_bytes()
                .to_vec(),
            BNC_COMMENT => comment_identifier
                .unwrap_or(0x99aa_bbcc)
                .to_le_bytes()
                .to_vec(),
            _ => flag.wrapping_mul(3).wrapping_add(7).to_le_bytes().to_vec(),
        };
        assert_eq!(field.len(), length, "BNC field layout is inconsistent");
        bytes.extend_from_slice(&field);
    }
    bytes
}

#[test]
fn every_pre_bnc_version_classifies_the_legacy_scalar_vocabulary() {
    for version in 0..=4 {
        assert_decoded(
            &pre_cell(version, 0, 0, None, None, None, None, None, None, None),
            ValueSource::Empty,
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                2,
                PRE_NUMBER,
                Some(12.5),
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ValueSource::Number(finite(12.5)),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                3,
                PRE_STRING,
                None,
                None,
                Some(17),
                None,
                None,
                None,
                None,
            ),
            ValueSource::Text(17),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                5,
                PRE_DATE,
                None,
                Some(345.25),
                None,
                None,
                None,
                None,
                None,
            ),
            ValueSource::Date(finite(345.25)),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                6,
                PRE_NUMBER,
                Some(2.0),
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ValueSource::Boolean(true),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                7,
                PRE_NUMBER,
                Some(-3.25),
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ValueSource::Duration(finite(-3.25)),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                8,
                PRE_FORMULA_ERROR,
                None,
                None,
                None,
                None,
                None,
                Some(23),
                None,
            ),
            ValueSource::Error(Some(23)),
            None,
        );
        assert_decoded(
            &pre_cell(
                version,
                9,
                PRE_RICH_TEXT,
                None,
                None,
                None,
                Some(29),
                None,
                None,
                None,
            ),
            ValueSource::RichText(29),
            None,
        );
    }
}

#[test]
fn every_bnc_v5_scalar_type_and_default_is_projected_without_sidecars() {
    assert_decoded(
        &bnc_cell(0, 0, None, None, None, None, None, None, None, None),
        ValueSource::Empty,
        None,
    );
    assert_decoded(
        &bnc_cell(
            2,
            BNC_DECIMAL,
            Some(decimal128(125, -1, false)),
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Number(finite(12.5)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            3,
            BNC_STRING,
            None,
            None,
            None,
            Some(17),
            None,
            None,
            None,
            None,
        ),
        ValueSource::Text(17),
        None,
    );
    assert_decoded(
        &bnc_cell(
            5,
            BNC_DATE,
            None,
            None,
            Some(345.25),
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Date(finite(345.25)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            6,
            BNC_NUMBER,
            None,
            Some(1.0),
            None,
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Boolean(true),
        None,
    );
    assert_decoded(
        &bnc_cell(
            7,
            BNC_NUMBER,
            None,
            Some(-3.25),
            None,
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Duration(finite(-3.25)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            8,
            BNC_FORMULA_ERROR,
            None,
            None,
            None,
            None,
            None,
            None,
            Some(23),
            None,
        ),
        ValueSource::Error(Some(23)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            9,
            BNC_RICH_TEXT,
            None,
            None,
            None,
            None,
            Some(29),
            None,
            None,
            None,
        ),
        ValueSource::RichText(29),
        None,
    );

    // BNC permits scalar-shaped cells with an omitted value field.  The
    // reader's typed defaults are part of the package compatibility contract.
    assert_decoded(
        &bnc_cell(2, 0, None, None, None, None, None, None, None, None),
        ValueSource::Number(finite(0.0)),
        None,
    );
    assert_decoded(
        &bnc_cell(5, 0, None, None, None, None, None, None, None, None),
        ValueSource::Date(finite(0.0)),
        None,
    );
    assert_decoded(
        &bnc_cell(6, 0, None, None, None, None, None, None, None, None),
        ValueSource::Boolean(false),
        None,
    );
    assert_decoded(
        &bnc_cell(7, 0, None, None, None, None, None, None, None, None),
        ValueSource::Duration(finite(0.0)),
        None,
    );
    assert_decoded(
        &bnc_cell(3, 0, None, None, None, None, None, None, None, None),
        ValueSource::Empty,
        None,
    );
    assert_decoded(
        &bnc_cell(9, 0, None, None, None, None, None, None, None, None),
        ValueSource::Number(finite(0.0)),
        None,
    );
    assert_decoded(
        &bnc_cell(8, 0, None, None, None, None, None, None, None, None),
        ValueSource::Error(None),
        None,
    );
}

#[test]
fn modern_numeric_types_nine_and_ten_are_distinct_from_legacy_rich_text_type_nine() {
    assert_decoded(
        &bnc_cell(
            9,
            BNC_NUMBER,
            None,
            Some(17.25),
            None,
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Number(finite(17.25)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            10,
            BNC_NUMBER,
            None,
            Some(-4.5),
            None,
            None,
            None,
            None,
            None,
            None,
        ),
        ValueSource::Number(finite(-4.5)),
        None,
    );
    assert_decoded(
        &bnc_cell(
            9,
            BNC_RICH_TEXT,
            None,
            Some(99.0),
            None,
            None,
            Some(41),
            None,
            None,
            None,
        ),
        ValueSource::RichText(41),
        None,
    );
    for version in 0..=4 {
        assert_decoded(
            &pre_cell(
                version,
                9,
                PRE_NUMBER,
                Some(99.0),
                None,
                None,
                None,
                None,
                None,
                None,
            ),
            ValueSource::Empty,
            None,
        );
    }
}

#[test]
fn identifiers_preserve_presence_including_zero_and_comments_are_orthogonal() {
    for version in 0..=4 {
        assert_decoded(
            &pre_cell(
                version,
                3,
                PRE_STRING | PRE_COMMENT,
                None,
                None,
                Some(0),
                None,
                None,
                None,
                Some(0),
            ),
            ValueSource::Text(0),
            Some(0),
        );
        assert_decoded(
            &pre_cell(
                version,
                3,
                PRE_COMMENT,
                None,
                None,
                None,
                None,
                None,
                None,
                Some(7),
            ),
            ValueSource::Empty,
            Some(7),
        );
        assert_decoded(
            &pre_cell(version, 9, 0, None, None, None, None, None, None, None),
            ValueSource::Empty,
            None,
        );
        assert_decoded(
            &pre_cell(version, 8, 0, None, None, None, None, None, None, None),
            ValueSource::Error(None),
            None,
        );
    }

    assert_decoded(
        &bnc_cell(
            3,
            BNC_STRING | BNC_COMMENT,
            None,
            None,
            None,
            Some(0),
            None,
            None,
            None,
            Some(0),
        ),
        ValueSource::Text(0),
        Some(0),
    );
    assert_decoded(
        &bnc_cell(
            9,
            BNC_RICH_TEXT | BNC_COMMENT,
            None,
            None,
            None,
            None,
            Some(0),
            None,
            None,
            Some(7),
        ),
        ValueSource::RichText(0),
        Some(7),
    );
    assert_decoded(
        &bnc_cell(
            8,
            BNC_COMMENT,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            Some(11),
        ),
        ValueSource::Error(None),
        Some(11),
    );
}

#[test]
fn formula_references_override_cached_values_but_retain_comment_references() {
    for version in 0..=4 {
        assert_decoded(
            &pre_cell(
                version,
                2,
                PRE_FORMULA | PRE_NUMBER | PRE_COMMENT,
                Some(88.0),
                None,
                None,
                None,
                Some(0),
                None,
                Some(3),
            ),
            ValueSource::Formula(0),
            Some(3),
        );
    }
    assert_decoded(
        &bnc_cell(
            2,
            BNC_FORMULA | BNC_DECIMAL | BNC_COMMENT,
            Some(decimal128(725, -2, false)),
            None,
            None,
            None,
            None,
            Some(0),
            None,
            Some(3),
        ),
        ValueSource::Formula(0),
        Some(3),
    );
    assert_decoded(
        &bnc_cell(
            6,
            BNC_FORMULA | BNC_NUMBER | BNC_COMMENT,
            None,
            Some(1.0),
            None,
            None,
            None,
            Some(41),
            None,
            Some(43),
        ),
        ValueSource::Formula(41),
        Some(43),
    );
}

#[test]
fn malformed_headers_truncated_fields_and_unknown_versions_fail_closed() {
    assert!(matches!(decode_cell_value(&[]), Err(DecodeError::Empty)));
    for version in 0..=4 {
        let source = pre_cell(
            version,
            2,
            PRE_NUMBER,
            Some(12.5),
            None,
            None,
            None,
            None,
            None,
            None,
        );
        for length in 1..source.len() {
            assert!(
                matches!(
                    decode_cell_value(&source[..length]),
                    Err(DecodeError::PreBnc(_))
                ),
                "pre-BNC v{version} accepted truncation at {length}"
            );
        }
    }
    let source = bnc_cell(
        2,
        BNC_DECIMAL,
        Some(decimal128(125, -1, false)),
        None,
        None,
        None,
        None,
        None,
        None,
        None,
    );
    for length in 1..source.len() {
        assert!(
            matches!(
                decode_cell_value(&source[..length]),
                Err(DecodeError::Bnc(_))
            ),
            "BNC accepted truncation at {length}"
        );
    }
    for version in [6, 7, u8::MAX] {
        assert!(matches!(
            decode_cell_value(&[version, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
            Err(DecodeError::UnsupportedVersion(actual)) if actual == version
        ));
    }
    assert!(matches!(
        decode_cell_value(&pre_cell(2, 1, 0, None, None, None, None, None, None, None)),
        Err(DecodeError::UnsupportedCellType(1))
    ));
    assert!(matches!(
        decode_cell_value(&bnc_cell(
            1, 0, None, None, None, None, None, None, None, None
        )),
        Err(DecodeError::UnsupportedCellType(1))
    ));
    assert!(matches!(
        decode_cell_value(&bnc_cell(
            0,
            0x8000_0000,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None
        )),
        Err(DecodeError::Bnc(_))
    ));
}

#[test]
fn non_finite_scalars_are_rejected_before_formula_precedence() {
    for version in 0..=4 {
        for (cell_type, scalar_flag, number, date) in [
            (2, PRE_NUMBER, Some(f64::NAN), None),
            (5, PRE_DATE, None, Some(f64::INFINITY)),
        ] {
            let source = pre_cell(
                version,
                cell_type,
                PRE_FORMULA | scalar_flag,
                number,
                date,
                None,
                None,
                Some(7),
                None,
                None,
            );
            assert!(matches!(
                decode_cell_value(&source),
                Err(DecodeError::PreBnc(_))
            ));
        }
    }

    for (cell_type, scalar_flag, number, date, decimal) in [
        (2, BNC_NUMBER, Some(f64::NAN), None, None),
        (5, BNC_DATE, None, Some(f64::NEG_INFINITY), None),
        (2, BNC_DECIMAL, None, None, Some(nonfinite_decimal128())),
    ] {
        let source = bnc_cell(
            cell_type,
            BNC_FORMULA | scalar_flag,
            decimal,
            number,
            date,
            None,
            None,
            Some(7),
            None,
            None,
        );
        assert!(matches!(
            decode_cell_value(&source),
            Err(DecodeError::Bnc(_))
        ));
    }
}

#[test]
fn projection_is_copy_drop_free_and_survives_source_storage() {
    fn assert_copy<T: Copy>() {}

    assert_copy::<CellValueSource>();
    assert_copy::<ValueSource>();
    assert!(!needs_drop::<CellValueSource>());
    assert!(!needs_drop::<ValueSource>());
    assert!(size_of::<CellValueSource>() <= 64);
    assert!(size_of::<ValueSource>() <= 64);

    let decoded = {
        let source = bnc_cell(
            2,
            BNC_DECIMAL | BNC_COMMENT,
            Some(decimal128(12345, -3, false)),
            None,
            None,
            None,
            None,
            None,
            None,
            Some(9),
        );
        decode_cell_value(&source).expect("cell source should decode")
    };
    assert_value(decoded.value, ValueSource::Number(finite(12.345)));
    assert_eq!(decoded.comment_identifier, Some(9));
}
