//! Regression coverage for the borrowed pre-BNC Numbers cell view.
//!
//! The fixtures in this file are assembled from the legacy fixed-layout wire
//! contract instead of going through the production encoder.  That keeps the
//! tests useful as an independent oracle while the pre-BNC reader is migrated
//! out of the package extractor.

use litchi_iwa_common::formula::FiniteF64;
use litchi_numbers_wire::pre_bnc::PreBncCellView;

const FLAG_RICH_TEXT: u32 = 0x0000_0200;
const FLAG_FORMULA: u32 = 0x0000_0008;
const FLAG_FORMULA_ERROR: u32 = 0x0000_0100;
const FLAG_STRING: u32 = 0x0000_0010;
const FLAG_NUMBER: u32 = 0x0000_0020;
const FLAG_DATE: u32 = 0x0000_0040;
const FLAG_COMMENT: u32 = 0x0000_1000;

const FIELD_LAYOUT: &[(u32, usize)] = &[
    (0x0000_0002, 4),
    (0x0000_0080, 4),
    (0x0000_0400, 4),
    (0x0000_0800, 4),
    (0x0000_0004, 4),
    (FLAG_FORMULA, 4),
    (FLAG_FORMULA_ERROR, 4),
    (FLAG_RICH_TEXT, 4),
    (FLAG_COMMENT, 4),
    (0x0000_2000, 4),
    (FLAG_STRING, 4),
    (FLAG_NUMBER, 8),
    (FLAG_DATE, 8),
    (0x0001_0000, 4),
    (0x0008_0000, 4),
    (0x0002_0000, 4),
    (0x0004_0000, 4),
    (0x0010_0000, 4),
    (0x0020_0000, 4),
    (0x0040_0000, 4),
    (0x0080_0000, 4),
];

const LOW_FLAGS: u32 = 0x0000_ffff;
const ALL_FLAGS: u32 = 0x00ff_3ffe;

fn header_length(version: u8) -> usize {
    if version <= 1 { 8 } else { 12 }
}

fn build_header(version: u8, cell_type: u8, flags: u32) -> Vec<u8> {
    let mut bytes = vec![0; header_length(version)];
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

fn field_bytes(flag: u32, length: usize) -> Vec<u8> {
    match flag {
        FLAG_NUMBER => 12.5_f64.to_le_bytes().to_vec(),
        FLAG_DATE => 345.25_f64.to_le_bytes().to_vec(),
        FLAG_FORMULA => 0x1020_3040_u32.to_le_bytes().to_vec(),
        FLAG_FORMULA_ERROR => 0x5060_7080_u32.to_le_bytes().to_vec(),
        FLAG_RICH_TEXT => 0x1122_3344_u32.to_le_bytes().to_vec(),
        FLAG_STRING => 0x5566_7788_u32.to_le_bytes().to_vec(),
        FLAG_COMMENT => 0x99aa_bbcc_u32.to_le_bytes().to_vec(),
        _ => {
            let value = flag.wrapping_mul(3).wrapping_add(7).to_le_bytes();
            value[..length].to_vec()
        },
    }
}

fn cell_with_flags(version: u8, cell_type: u8, flags: u32, tail: &[u8]) -> Vec<u8> {
    let mut bytes = build_header(version, cell_type, flags);
    for &(flag, length) in FIELD_LAYOUT {
        if flags & flag != 0 {
            let value = field_bytes(flag, length);
            assert_eq!(value.len(), length, "test field layout is inconsistent");
            bytes.extend_from_slice(&value);
        }
    }
    bytes.extend_from_slice(tail);
    bytes
}

fn cell_with_zero_ids(version: u8) -> Vec<u8> {
    let flags = FLAG_STRING | FLAG_RICH_TEXT | FLAG_FORMULA | FLAG_FORMULA_ERROR | FLAG_COMMENT;
    let mut bytes = build_header(version, 9, flags);
    for &(flag, length) in FIELD_LAYOUT {
        if flags & flag != 0 {
            assert_eq!(length, 4);
            bytes.extend_from_slice(&0_u32.to_le_bytes());
        }
    }
    bytes
}

fn assert_finite(value: Option<FiniteF64>, expected: f64) {
    assert_eq!(value.map(FiniteF64::get), Some(expected));
}

#[test]
fn every_legacy_version_exposes_all_fields_from_the_fixed_order() {
    for version in 0..=4 {
        let flags = if version <= 1 {
            ALL_FLAGS & LOW_FLAGS
        } else {
            ALL_FLAGS
        };
        let source = cell_with_flags(version, 9, flags, &[]);
        let view = PreBncCellView::parse(&source)
            .unwrap_or_else(|_| panic!("version {version} did not parse"));

        assert_eq!(view.version(), version);
        assert_eq!(view.source(), source.as_slice());
        assert_eq!(view.cell_type(), 9);
        assert_finite(view.number(), 12.5);
        assert_finite(view.date(), 345.25);
        assert_eq!(view.string_identifier(), Some(0x5566_7788));
        assert_eq!(view.rich_text_identifier(), Some(0x1122_3344));
        assert_eq!(view.formula_identifier(), Some(0x1020_3040));
        assert_eq!(view.formula_error_identifier(), Some(0x5060_7080));
        assert_eq!(view.comment_identifier(), Some(0x99aa_bbcc));
    }
}

#[test]
fn zero_identifiers_remain_present_in_all_legacy_versions() {
    for version in 0..=4 {
        let source = cell_with_zero_ids(version);
        let view = PreBncCellView::parse(&source)
            .unwrap_or_else(|_| panic!("version {version} did not parse"));
        assert_eq!(view.string_identifier(), Some(0));
        assert_eq!(view.rich_text_identifier(), Some(0));
        assert_eq!(view.formula_identifier(), Some(0));
        assert_eq!(view.formula_error_identifier(), Some(0));
        assert_eq!(view.comment_identifier(), Some(0));
    }
}

#[test]
fn every_truncated_prefix_before_the_known_fields_is_rejected() {
    for version in 0..=4 {
        let flags = if version <= 1 {
            ALL_FLAGS & LOW_FLAGS
        } else {
            ALL_FLAGS
        };
        let source = cell_with_flags(version, 9, flags, &[]);
        let known_end = source.len();
        for length in 0..known_end {
            assert!(
                PreBncCellView::parse(&source[..length]).is_err(),
                "version {version} truncation at {length} was accepted"
            );
        }
    }
}

#[test]
fn non_finite_scalars_are_rejected_even_when_a_formula_identifier_wins() {
    for version in 0..=4 {
        for raw in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            for scalar_flag in [FLAG_NUMBER, FLAG_DATE] {
                let flags = FLAG_FORMULA | scalar_flag;
                let mut source = build_header(version, 0, flags);
                for &(flag, length) in FIELD_LAYOUT {
                    if flags & flag == 0 {
                        continue;
                    }
                    if flag == scalar_flag {
                        source.extend_from_slice(&raw.to_le_bytes());
                    } else {
                        source.extend_from_slice(&field_bytes(flag, length));
                    }
                }
                assert!(
                    PreBncCellView::parse(&source).is_err(),
                    "version {version}, flag 0x{scalar_flag:08x}, value {raw:?} bypassed finite validation"
                );
            }
        }
    }
}

#[test]
fn unknown_flags_and_opaque_tail_are_preserved_without_copying_source() {
    for version in [0, 1, 2, 3, 4] {
        let unknown_flag = if version <= 1 { 0x8000 } else { 0x8000_0000 };
        let known_flags = FLAG_FORMULA | FLAG_STRING;
        let unknown_payload = [0xde, 0xad, 0xbe, 0xef];
        let explicit_tail = [0xca, 0xfe, 0xba, 0xbe];
        let mut source = cell_with_flags(version, 3, known_flags | unknown_flag, &unknown_payload);
        source.extend_from_slice(&explicit_tail);

        let view = PreBncCellView::parse(&source)
            .unwrap_or_else(|_| panic!("version {version} did not parse"));
        assert_eq!(view.source(), source.as_slice());
        assert_eq!(view.cell_type(), 3);
        assert_eq!(view.formula_identifier(), Some(0x1020_3040));
        assert_eq!(view.string_identifier(), Some(0x5566_7788));

        let expected_tail = [&unknown_payload[..], &explicit_tail[..]].concat();
        assert_eq!(view.opaque_tail(), expected_tail.as_slice());
        if !source.is_empty() {
            assert_eq!(view.source().as_ptr(), source.as_ptr());
            assert_eq!(
                view.opaque_tail().as_ptr(),
                source[view.source().len() - view.opaque_tail().len()..].as_ptr()
            );
        }
    }
}

#[test]
fn modern_and_future_versions_are_rejected_by_the_legacy_reader() {
    for version in [5, 6, 0xff] {
        let source = cell_with_flags(version, 3, FLAG_STRING, &[]);
        assert!(
            PreBncCellView::parse(&source).is_err(),
            "version {version} was accepted as pre-BNC"
        );
    }
}
