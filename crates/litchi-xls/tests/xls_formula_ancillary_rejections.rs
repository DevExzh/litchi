//! Independent wire-level checks for the public Formula record reader.

use litchi_xls::records::{CellRecord, Encoding};

fn payload(tokens: &[u8], extra: &[u8]) -> Vec<u8> {
    let mut bytes = vec![0; 22];
    bytes[20..22].copy_from_slice(&u16::try_from(tokens.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(tokens);
    bytes.extend_from_slice(extra);
    bytes
}

fn parse(bytes: &[u8]) -> litchi_xls::Result<CellRecord> {
    CellRecord::parse(0x0006, bytes, &Encoding::Utf16Le)
}

fn memory_tokens() -> Vec<u8> {
    // PtgMemArea, two PtgRef operands, and their rectangular range. The four unused
    // bytes are intentionally nonzero and must survive without normalization.
    vec![
        0x46, 0x12, 0x34, 0x56, 0x78, 11, 0, 0x24, 0, 0, 0, 0, 0x24, 1, 0, 1, 0, 0x11,
    ]
}

fn memory_extra() -> Vec<u8> {
    // One Ref8U: rows 0..=1, columns 0..=1.
    vec![1, 0, 0, 0, 1, 0, 0, 0, 1, 0]
}

fn array_extra() -> Vec<u8> {
    // One column, one row, and one SerNum value.
    let mut bytes = vec![0, 0, 0, 1];
    bytes.extend_from_slice(&1.0_f64.to_le_bytes());
    bytes
}

#[test]
fn mixed_extra_structures_follow_token_order_and_consume_the_exact_tail() {
    let mut tokens = vec![0x40, 0, 0, 0, 0, 0, 0, 0];
    tokens.extend_from_slice(&memory_tokens());
    tokens.push(0x03);
    let mut extra = array_extra();
    extra.extend_from_slice(&memory_extra());
    let CellRecord::Formula {
        formula, metadata, ..
    } = parse(&payload(&tokens, &extra)).unwrap()
    else {
        panic!("expected a Formula record");
    };
    assert_eq!(formula, tokens);
    assert_eq!(metadata.ancillary_bytes(), Some(extra.as_slice()));

    for length in 1..extra.len() {
        assert!(
            parse(&payload(&tokens, &extra[..length])).is_err(),
            "accepted truncated ancillary tail at byte {length}"
        );
    }
    let mut reordered = memory_extra();
    reordered.extend_from_slice(&array_extra());
    assert!(parse(&payload(&tokens, &reordered)).is_err());
    extra.push(0);
    assert!(parse(&payload(&tokens, &extra)).is_err());
}

#[test]
fn memory_ranges_and_token_boundaries_are_checked_before_retention() {
    let tokens = memory_tokens();
    let extra = memory_extra();
    assert!(parse(&payload(&tokens, &extra)).is_ok());

    for (offset, value) in [(0, u16::MAX), (2, 2), (6, 2), (8, 256)] {
        let mut malformed = extra.clone();
        malformed[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
        assert!(
            parse(&payload(&tokens, &malformed)).is_err(),
            "accepted invalid range field at {offset}"
        );
    }
    for count in [0_u16, 1, u16::MAX] {
        let mut malformed = tokens.clone();
        malformed[5..7].copy_from_slice(&count.to_le_bytes());
        assert!(parse(&payload(&malformed, &extra)).is_err());
    }
    let mut high_bit = tokens.clone();
    high_bit[0] |= 0x80;
    assert!(parse(&payload(&high_bit, &extra)).is_err());

    for length in 0..22 {
        assert!(parse(&payload(&tokens, &extra)[..length]).is_err());
    }
    let mut overrun = payload(&tokens, &extra);
    overrun[20..22].copy_from_slice(&u16::MAX.to_le_bytes());
    assert!(parse(&overrun).is_err());
    let mut reserved_flags = payload(&tokens, &extra);
    reserved_flags[14] = 0x02;
    assert!(parse(&reserved_flags).is_err());
}

#[test]
fn array_extra_values_keep_the_existing_strict_scalar_checks() {
    let tokens = [0x40, 0, 0, 0, 0, 0, 0, 0];
    assert!(parse(&payload(&tokens, &array_extra())).is_ok());
    let mut boolean = vec![0, 0, 0, 4, 2, 0, 0, 0, 0, 0, 0, 0];
    assert!(matches!(
        parse(&payload(&tokens, &boolean)),
        Err(litchi_xls::Error::InvalidRecord {
            record_type: 0x0006,
            ..
        })
    ));
    boolean[4] = 1;
    boolean[5] = 1;
    assert!(parse(&payload(&tokens, &boolean)).is_err());
    // One UTF-16 code unit containing an unpaired high surrogate.
    assert!(parse(&payload(&tokens, &[0, 0, 0, 2, 1, 0, 1, 0, 0xd8])).is_err());
    let mut oversized = array_extra();
    oversized[..3].fill(0xff);
    assert!(parse(&payload(&tokens, &oversized)).is_err());
    let mut nonfinite = array_extra();
    nonfinite[4..].copy_from_slice(&f64::INFINITY.to_le_bytes());
    assert!(parse(&payload(&tokens, &nonfinite)).is_err());
}

#[test]
fn suffix_support_does_not_redefine_legacy_opaque_token_validation() {
    // Empty rgcb retains the pre-existing opaque-token contract. This reader
    // is not a complete RPN validator and does not certify missing extras.
    for tokens in [vec![0xff], memory_tokens()] {
        let CellRecord::Formula { metadata, .. } = parse(&payload(&tokens, &[])).unwrap() else {
            panic!("expected a Formula record");
        };
        assert!(metadata.ancillary_bytes().is_none());
    }
    assert!(parse(&payload(&[0xff], &[0])).is_err());
    assert!(parse(&payload(&[0x1e, 1, 0], &memory_extra())).is_err());
}

#[test]
fn ordinary_name_and_three_dimensional_references_do_not_require_revision_extras() {
    // CellParsedFormula is outside revision context. These operands do not own
    // an extra there; the following PtgArray owns the sole PtgExtraArray.
    for (opcode, length) in [
        (0x43, 5),
        (0x59, 7),
        (0x5a, 7),
        (0x5b, 11),
        (0x5c, 7),
        (0x5d, 11),
    ] {
        let mut tokens = vec![0; length];
        tokens[0] = opcode;
        if opcode == 0x43 {
            tokens[1] = 1;
        } else if opcode == 0x59 {
            tokens[3] = 1;
        }
        tokens.extend_from_slice(&[0x40, 0, 0, 0, 0, 0, 0, 0, 0x03]);
        let extra = array_extra();
        let CellRecord::Formula {
            formula, metadata, ..
        } = parse(&payload(&tokens, &extra)).unwrap()
        else {
            panic!("expected a Formula record");
        };
        assert_eq!(formula, tokens);
        assert_eq!(metadata.ancillary_bytes(), Some(extra.as_slice()));
    }
}

#[test]
fn nonempty_array_extra_requires_a_value_or_array_typed_token() {
    for opcode in [0x40, 0x60] {
        assert!(parse(&payload(&[opcode, 0, 0, 0, 0, 0, 0, 0], &array_extra())).is_ok());
    }
    assert!(parse(&payload(&[0x20, 0, 0, 0, 0, 0, 0, 0], &array_extra())).is_err());
}

#[test]
fn the_record_limit_includes_both_tokens_and_ancillary_bytes() {
    let tokens = [0x40, 0, 0, 0, 0, 0, 0, 0];
    // One column and 910 rows: one ten-byte SerStr and 909 nine-byte SerNums
    // place the complete Formula payload exactly at the 8,224-byte limit.
    let mut extra = vec![0];
    extra.extend_from_slice(&909_u16.to_le_bytes());
    extra.extend_from_slice(&[2, 6, 0, 0]);
    extra.extend_from_slice(b"scalar");
    for _ in 0..909 {
        extra.push(1);
        extra.extend_from_slice(&1.0_f64.to_le_bytes());
    }
    let mut bytes = payload(&tokens, &extra);
    assert_eq!(bytes.len(), 8_224);
    assert!(parse(&bytes).is_ok());
    bytes.push(0);
    assert!(parse(&bytes).is_err());
}
