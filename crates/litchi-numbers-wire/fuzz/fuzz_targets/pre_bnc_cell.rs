#![no_main]

use std::{hint::black_box, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_numbers_wire::pre_bnc::PreBncCellView;

const MAX_INPUT_BYTES: usize = 64 * 1024;

const FLAG_UNKNOWN_0002: u32 = 0x0000_0002;
const FLAG_UNKNOWN_0080: u32 = 0x0000_0080;
const FLAG_UNKNOWN_0400: u32 = 0x0000_0400;
const FLAG_UNKNOWN_0800: u32 = 0x0000_0800;
const FLAG_UNKNOWN_0004: u32 = 0x0000_0004;
const FLAG_FORMULA: u32 = 0x0000_0008;
const FLAG_FORMULA_ERROR: u32 = 0x0000_0100;
const FLAG_RICH_TEXT: u32 = 0x0000_0200;
const FLAG_COMMENT: u32 = 0x0000_1000;
const FLAG_UNKNOWN_2000: u32 = 0x0000_2000;
const FLAG_STRING: u32 = 0x0000_0010;
const FLAG_NUMBER: u32 = 0x0000_0020;
const FLAG_DATE: u32 = 0x0000_0040;
const FLAG_UNKNOWN_010000: u32 = 0x0001_0000;
const FLAG_UNKNOWN_080000: u32 = 0x0008_0000;
const FLAG_UNKNOWN_020000: u32 = 0x0002_0000;
const FLAG_UNKNOWN_040000: u32 = 0x0004_0000;
const FLAG_UNKNOWN_100000: u32 = 0x0010_0000;
const FLAG_UNKNOWN_200000: u32 = 0x0020_0000;
const FLAG_UNKNOWN_400000: u32 = 0x0040_0000;
const FLAG_UNKNOWN_800000: u32 = 0x0080_0000;

const FIELD_LAYOUT: &[(u32, usize)] = &[
    (FLAG_UNKNOWN_0002, 4),
    (FLAG_UNKNOWN_0080, 4),
    (FLAG_UNKNOWN_0400, 4),
    (FLAG_UNKNOWN_0800, 4),
    (FLAG_UNKNOWN_0004, 4),
    (FLAG_FORMULA, 4),
    (FLAG_FORMULA_ERROR, 4),
    (FLAG_RICH_TEXT, 4),
    (FLAG_COMMENT, 4),
    (FLAG_UNKNOWN_2000, 4),
    (FLAG_STRING, 4),
    (FLAG_NUMBER, 8),
    (FLAG_DATE, 8),
    (FLAG_UNKNOWN_010000, 4),
    (FLAG_UNKNOWN_080000, 4),
    (FLAG_UNKNOWN_020000, 4),
    (FLAG_UNKNOWN_040000, 4),
    (FLAG_UNKNOWN_100000, 4),
    (FLAG_UNKNOWN_200000, 4),
    (FLAG_UNKNOWN_400000, 4),
    (FLAG_UNKNOWN_800000, 4),
];

const LOW_FLAGS: u32 = 0x0000_ffff;
const ALL_FLAGS: u32 = 0x00ff_3ffe;

static FIXED_CASES: OnceLock<()> = OnceLock::new();

fuzz_target!(|data: &[u8]| {
    if data.len() <= MAX_INPUT_BYTES {
        exercise_source(data);
        exercise_mutated_versions(data);

        if let Some(decoded) = decode_hex_input(data) {
            exercise_source(&decoded);
            exercise_mutated_versions(&decoded);
        }
    }

    FIXED_CASES.get_or_init(run_fixed_cases);
});

fn exercise_source(source: &[u8]) {
    let snapshot = source.to_vec();

    match PreBncCellView::parse(source) {
        Ok(view) => {
            assert_eq!(view.source(), source);
            assert_eq!(view.source().as_ptr(), source.as_ptr());
            assert_eq!(view.source().len(), source.len());
            assert_borrowed(source, view.source());
            assert_borrowed(source, view.opaque_tail());

            let tail = view.opaque_tail();
            if !tail.is_empty() {
                let suffix = &source[source.len() - tail.len()..];
                assert_eq!(tail, suffix);
                assert_eq!(tail.as_ptr(), suffix.as_ptr());
            }

            if view.flags() & FLAG_NUMBER != 0 {
                assert!(view.number().is_some());
            }
            if view.flags() & FLAG_DATE != 0 {
                assert!(view.date().is_some());
            }
            if view.flags() & FLAG_STRING != 0 {
                assert!(view.string_identifier().is_some());
            }
            if view.flags() & FLAG_RICH_TEXT != 0 {
                assert!(view.rich_text_identifier().is_some());
            }
            if view.flags() & FLAG_FORMULA != 0 {
                assert!(view.formula_identifier().is_some());
            }
            if view.flags() & FLAG_FORMULA_ERROR != 0 {
                assert!(view.formula_error_identifier().is_some());
            }
            if view.flags() & FLAG_COMMENT != 0 {
                assert!(view.comment_identifier().is_some());
            }

            if let Some(number) = view.number() {
                assert!(number.get().is_finite());
            }
            if let Some(date) = view.date() {
                assert!(date.get().is_finite());
            }

            black_box((
                view.version(),
                view.cell_type(),
                view.flags(),
                view.number().map(|value| value.get()),
                view.date().map(|value| value.get()),
                view.string_identifier(),
                view.rich_text_identifier(),
                view.formula_identifier(),
                view.formula_error_identifier(),
                view.comment_identifier(),
                view.opaque_tail().len(),
            ));
        },
        Err(error) => {
            // Malformed input is an expected fuzzer path. Formatting the
            // error also keeps the error value live without assuming a
            // particular error variant or message.
            black_box(error.to_string());
        },
    }

    assert_eq!(source, snapshot.as_slice());
}

fn assert_borrowed(source: &[u8], borrowed: &[u8]) {
    if borrowed.is_empty() {
        return;
    }

    let source_start = source.as_ptr() as usize;
    let source_end = source_start
        .checked_add(source.len())
        .expect("source pointer range overflow");
    let borrowed_start = borrowed.as_ptr() as usize;
    let borrowed_end = borrowed_start
        .checked_add(borrowed.len())
        .expect("borrowed pointer range overflow");

    assert!(borrowed_start >= source_start);
    assert!(borrowed_end <= source_end);
}

fn exercise_mutated_versions(data: &[u8]) {
    let mut entropy_cursor = 0;
    let base_flags = next_u32(data, &mut entropy_cursor, FLAG_FORMULA | FLAG_STRING) & ALL_FLAGS;

    for version in 0..=4_u8 {
        let mut flags = base_flags.rotate_left(u32::from(version) * 7) & ALL_FLAGS;
        if flags == 0 {
            flags = FLAG_FORMULA | FLAG_STRING;
        }
        if version <= 1 {
            flags &= LOW_FLAGS;
            if flags == 0 {
                flags = FLAG_FORMULA | FLAG_STRING;
            }
        }

        let cell_type = data
            .get(usize::from(version))
            .copied()
            .unwrap_or(version.wrapping_mul(37));
        let tail = input_tail(data, usize::from(version) + 1);
        let source = cell_with_flags(version, cell_type, flags, data, &tail);
        exercise_source(&source);

        // Ensure each version also sees a sparse layout mutation even when
        // the fuzzer's flag word happens to be zero or repeats one bit.
        let field_count = if version <= 1 { 13 } else { FIELD_LAYOUT.len() };
        let field_index = (usize::from(version) * 3) % field_count;
        let single_flag = FIELD_LAYOUT[field_index].0;
        let single_source = cell_with_flags(version, cell_type, single_flag, data, &tail);
        exercise_source(&single_source);
    }
}

fn run_fixed_cases() {
    let all_semantic_flags = FLAG_FORMULA
        | FLAG_FORMULA_ERROR
        | FLAG_RICH_TEXT
        | FLAG_COMMENT
        | FLAG_STRING
        | FLAG_NUMBER
        | FLAG_DATE;

    for version in 0..=4_u8 {
        let flags = if version <= 1 {
            ALL_FLAGS & LOW_FLAGS
        } else {
            ALL_FLAGS
        };
        let source = cell_with_flags(version, 9, flags, &[], &[0xca, 0xfe]);
        let view = PreBncCellView::parse(&source)
            .unwrap_or_else(|error| panic!("fixed version {version} case rejected: {error}"));
        assert_eq!(view.version(), version);
        assert_eq!(view.cell_type(), 9);
        assert_eq!(view.number().map(|value| value.get()), Some(12.5));
        assert_eq!(view.date().map(|value| value.get()), Some(345.25));
        assert_eq!(view.string_identifier(), Some(0x5566_7788));
        assert_eq!(view.rich_text_identifier(), Some(0x1122_3344));
        assert_eq!(view.formula_identifier(), Some(0x1020_3040));
        assert_eq!(view.formula_error_identifier(), Some(0x5060_7080));
        assert_eq!(view.comment_identifier(), Some(0x99aa_bbcc));
        exercise_source(&source);

        let semantic_source = cell_with_flags(
            version,
            3,
            if version <= 1 {
                all_semantic_flags & LOW_FLAGS
            } else {
                all_semantic_flags
            },
            &[],
            &[],
        );
        exercise_source(&semantic_source);

        let zero_source = cell_with_flags(
            version,
            7,
            FLAG_STRING | FLAG_RICH_TEXT | FLAG_FORMULA | FLAG_FORMULA_ERROR | FLAG_COMMENT,
            &[0; 20],
            &[],
        );
        let zero_view = PreBncCellView::parse(&zero_source)
            .unwrap_or_else(|error| panic!("fixed zero-ID version {version} rejected: {error}"));
        assert_eq!(zero_view.string_identifier(), Some(0));
        assert_eq!(zero_view.rich_text_identifier(), Some(0));
        assert_eq!(zero_view.formula_identifier(), Some(0));
        assert_eq!(zero_view.formula_error_identifier(), Some(0));
        assert_eq!(zero_view.comment_identifier(), Some(0));
    }

    for version in 0..=4_u8 {
        let unknown_flag = if version <= 1 {
            0x0000_8000
        } else {
            0x8000_0000
        };
        let known_flags = FLAG_FORMULA | FLAG_STRING;
        let unknown_payload = [0xde, 0xad, 0xbe, 0xef];
        let explicit_tail = [0xca, 0xfe, 0xba, 0xbe];
        let mut source = cell_with_flags(
            version,
            3,
            known_flags | unknown_flag,
            &[],
            &unknown_payload,
        );
        source.extend_from_slice(&explicit_tail);
        let view = PreBncCellView::parse(&source)
            .unwrap_or_else(|error| panic!("unknown flag version {version} rejected: {error}"));
        let expected_tail = [&unknown_payload[..], &explicit_tail[..]].concat();
        assert_eq!(view.opaque_tail(), expected_tail.as_slice());
        exercise_source(&source);
    }

    for version in 0..=4_u8 {
        let flags = if version <= 1 {
            ALL_FLAGS & LOW_FLAGS
        } else {
            ALL_FLAGS
        };
        let source = cell_with_flags(version, 9, flags, &[], &[]);
        for length in 0..source.len() {
            assert!(
                PreBncCellView::parse(&source[..length]).is_err(),
                "fixed version {version} truncation at {length} was accepted"
            );
        }
    }

    for version in 0..=4_u8 {
        for raw in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            for scalar_flag in [FLAG_NUMBER, FLAG_DATE] {
                let flags = FLAG_FORMULA | scalar_flag;
                let mut source = cell_with_flags(version, 0, flags, &[], &[]);
                let offset = field_offset(version, flags, scalar_flag)
                    .expect("scalar flag must be present in its fixed layout");
                source[offset..offset + 8].copy_from_slice(&raw.to_le_bytes());
                assert!(
                    PreBncCellView::parse(&source).is_err(),
                    "fixed version {version} accepted non-finite {scalar_flag:#x}"
                );
                exercise_source(&source);
            }
        }
    }

    for version in [5_u8, 6, u8::MAX] {
        let source = cell_with_flags(version, 3, FLAG_STRING, &[], &[]);
        assert!(
            PreBncCellView::parse(&source).is_err(),
            "future version {version} was accepted by the pre-BNC reader"
        );
    }
}

fn cell_with_flags(version: u8, cell_type: u8, flags: u32, entropy: &[u8], tail: &[u8]) -> Vec<u8> {
    let mut bytes = build_header(version, cell_type, flags);
    let mut cursor = 0;
    for &(flag, length) in FIELD_LAYOUT {
        if flags & flag != 0 {
            append_field(&mut bytes, flag, length, entropy, &mut cursor);
        }
    }
    bytes.extend_from_slice(tail);
    bytes
}

fn build_header(version: u8, cell_type: u8, flags: u32) -> Vec<u8> {
    let mut bytes = vec![0; if version <= 1 { 8 } else { 12 }];
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

fn append_field(bytes: &mut Vec<u8>, flag: u32, length: usize, entropy: &[u8], cursor: &mut usize) {
    match flag {
        FLAG_NUMBER => {
            let raw = next_u64(entropy, cursor, 12.5_f64.to_bits());
            let finite = finite_from_bits(raw, 12.5);
            bytes.extend_from_slice(&finite.to_le_bytes());
        },
        FLAG_DATE => {
            let raw = next_u64(entropy, cursor, 345.25_f64.to_bits());
            let finite = finite_from_bits(raw, 345.25);
            bytes.extend_from_slice(&finite.to_le_bytes());
        },
        FLAG_FORMULA => {
            bytes.extend_from_slice(&next_u32(entropy, cursor, 0x1020_3040_u32).to_le_bytes())
        },
        FLAG_FORMULA_ERROR => {
            bytes.extend_from_slice(&next_u32(entropy, cursor, 0x5060_7080_u32).to_le_bytes())
        },
        FLAG_RICH_TEXT => {
            bytes.extend_from_slice(&next_u32(entropy, cursor, 0x1122_3344_u32).to_le_bytes())
        },
        FLAG_STRING => {
            bytes.extend_from_slice(&next_u32(entropy, cursor, 0x5566_7788_u32).to_le_bytes())
        },
        FLAG_COMMENT => {
            bytes.extend_from_slice(&next_u32(entropy, cursor, 0x99aa_bbcc_u32).to_le_bytes())
        },
        _ => {
            let value = next_u32(entropy, cursor, flag.wrapping_mul(3).wrapping_add(7));
            bytes.extend_from_slice(&value.to_le_bytes()[..length]);
        },
    }
}

fn finite_from_bits(raw: u64, fallback: f64) -> f64 {
    let value = f64::from_bits(raw);
    if value.is_finite() { value } else { fallback }
}

fn field_offset(version: u8, flags: u32, target: u32) -> Option<usize> {
    let mut offset = if version <= 1 { 8 } else { 12 };
    for &(flag, length) in FIELD_LAYOUT {
        if flags & flag == 0 {
            continue;
        }
        if flag == target {
            return Some(offset);
        }
        offset += length;
    }
    None
}

fn input_tail(data: &[u8], start: usize) -> Vec<u8> {
    let length = data.get(start).map_or(0, |byte| usize::from(*byte) % 33);
    let start = start.saturating_add(1).min(data.len());
    let end = start.saturating_add(length).min(data.len());
    data[start..end].to_vec()
}

fn next_u32(data: &[u8], cursor: &mut usize, fallback: u32) -> u32 {
    let end = (*cursor).saturating_add(4);
    let value = if let Some(bytes) = data.get(*cursor..end) {
        if bytes.len() == 4 {
            u32::from_le_bytes(bytes.try_into().expect("checked four-byte slice"))
        } else {
            fallback
        }
    } else {
        fallback
    };
    *cursor = end;
    value
}

fn next_u64(data: &[u8], cursor: &mut usize, fallback: u64) -> u64 {
    let end = (*cursor).saturating_add(8);
    let value = if let Some(bytes) = data.get(*cursor..end) {
        if bytes.len() == 8 {
            u64::from_le_bytes(bytes.try_into().expect("checked eight-byte slice"))
        } else {
            fallback
        }
    } else {
        fallback
    };
    *cursor = end;
    value
}

fn decode_hex_input(data: &[u8]) -> Option<Vec<u8>> {
    if !data.starts_with(b"hex:") {
        return None;
    }

    let mut decoded = Vec::with_capacity((data.len() - 4) / 2);
    let mut high_nibble = None;
    for byte in data[4..].iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = match byte {
            b'0'..=b'9' => byte - b'0',
            b'a'..=b'f' => byte - b'a' + 10,
            b'A'..=b'F' => byte - b'A' + 10,
            _ => return None,
        };
        if let Some(high) = high_nibble.take() {
            if decoded.len() >= MAX_INPUT_BYTES {
                return None;
            }
            decoded.push((high << 4) | nibble);
        } else {
            high_nibble = Some(nibble);
        }
    }

    if high_nibble.is_some() {
        return None;
    }
    Some(decoded)
}
