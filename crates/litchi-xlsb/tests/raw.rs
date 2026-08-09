#![allow(
    clippy::pedantic,
    clippy::expect_used,
    clippy::map_err_ignore,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration tests use panic-on-failure extraction and exact wire fixture comparisons"
)]

use std::io::{Cursor as IoCursor, Read};

use litchi_xlsb::raw::{Cursor, Error, Header, Kind, Limits, Records, Stage, Writer, kind};

#[test]
fn round_trips_kind_boundaries_with_borrowed_payloads() {
    let cases = [
        (Kind::new(0x7f).unwrap(), vec![0x7f, 0x03, 1, 2, 3]),
        (Kind::new(0x80).unwrap(), vec![0x80, 0x01, 0x03, 1, 2, 3]),
        (
            Kind::new(Kind::MAX).unwrap(),
            vec![0xff, 0x7f, 0x03, 1, 2, 3],
        ),
    ];

    for (kind, expected) in cases {
        let mut bytes = Vec::new();
        Writer::new(&mut bytes)
            .write_record(kind, &[1, 2, 3])
            .unwrap();
        assert_eq!(bytes, expected);

        let record = Records::new(&bytes).next().unwrap().unwrap();
        assert_eq!(record.kind(), kind);
        assert_eq!(record.len(), 3);
        assert_eq!(record.payload(), [1, 2, 3]);
        assert_eq!(record.payload().as_ptr(), bytes[bytes.len() - 3..].as_ptr());
    }
}

#[test]
fn rejects_kinds_outside_the_two_byte_domain() {
    assert!(matches!(
        Kind::new(0x4000),
        Err(Error::KindOutOfRange { value: 0x4000 })
    ));
    assert!(matches!(
        Header::parse(&[0x80, 0x80, 0x00], Limits::DEFAULT),
        Err(Error::InvalidKind { offset: 0 })
    ));
    assert!(matches!(
        Header::parse(&[0x80, 0x00, 0x00], Limits::DEFAULT),
        Err(Error::InvalidKind { offset: 0 })
    ));
}

#[test]
fn fourth_length_continuation_bit_does_not_consume_a_fifth_byte() {
    let bytes = [0x00, 0x80, 0x80, 0x80, 0x80, 0x55];
    let mut input = IoCursor::new(bytes);
    let header = Header::read(&mut input, Limits::DEFAULT).unwrap().unwrap();
    assert_eq!(header.len(), 0);

    let mut next = [0_u8; 1];
    input.read_exact(&mut next).unwrap();
    assert_eq!(next, [0x55]);
}

#[test]
fn accepts_the_largest_four_byte_wire_length_when_budgeted() {
    let bytes = [0x00, 0xff, 0xff, 0xff, 0xff];
    let limits = Limits::new(0x0fff_ffff, 1);
    let (header, consumed) = Header::parse(&bytes, limits).unwrap();
    assert_eq!(header.len(), 0x0fff_ffff);
    assert_eq!(consumed, bytes.len());
}

#[test]
fn distinguishes_clean_eof_from_truncated_header_and_payload() {
    assert!(
        Header::read(&mut IoCursor::new([]), Limits::DEFAULT)
            .unwrap()
            .is_none()
    );
    assert!(matches!(
        Header::read(&mut IoCursor::new([0x80]), Limits::DEFAULT),
        Err(Error::Truncated {
            stage: Stage::Kind,
            ..
        })
    ));
    assert!(matches!(
        Header::read(&mut IoCursor::new([0x01]), Limits::DEFAULT),
        Err(Error::Truncated {
            stage: Stage::Length,
            offset: 1,
            ..
        })
    ));
    assert!(matches!(
        Header::read(&mut IoCursor::new([0x80, 0x01]), Limits::DEFAULT),
        Err(Error::Truncated {
            stage: Stage::Length,
            offset: 2,
            ..
        })
    ));

    let mut records = Records::new(&[0x01, 0x02, 0xaa]);
    assert!(matches!(
        records.next(),
        Some(Err(Error::Truncated {
            stage: Stage::Payload,
            needed: 2,
            available: 1,
            ..
        }))
    ));
    assert!(records.next().is_none());
}

#[test]
fn enforces_payload_budgets_before_reading_or_writing_payloads() {
    let limits = Limits::new(3, 10);
    assert!(matches!(
        Header::parse(&[0x01, 0x04], limits),
        Err(Error::PayloadLimit {
            length: 4,
            limit: 3,
            ..
        })
    ));

    let mut writer = Writer::with_limits(Vec::new(), limits);
    assert!(matches!(
        writer.write_header(kind::CELL_BLANK, 4),
        Err(Error::PayloadLimit {
            length: 4,
            limit: 3,
            ..
        })
    ));
}

#[test]
fn explicit_limits_are_not_silently_clamped_to_wire_limits() {
    let requested = 0x0fff_ffff_usize + 1;
    assert_eq!(Limits::new(requested, 7).payload(), requested);
}

#[test]
fn oversized_wire_headers_are_rejected_before_output() {
    let too_large = 0x0fff_ffff_usize + 1;
    let mut output = Vec::new();
    let mut writer = Writer::with_limits(&mut output, Limits::new(too_large, 1));
    assert!(matches!(
        writer.write_header(kind::CELL_BLANK, too_large),
        Err(Error::LengthOverflow {
            what: "record payload",
            length,
        }) if length == too_large
    ));
    assert!(output.is_empty());
}

#[test]
fn strictly_decodes_utf16_and_enforces_string_budgets() {
    let mut good = 2_u32.to_le_bytes().to_vec();
    good.extend_from_slice(&[b'A', 0, b'B', 0]);
    let mut cursor = Cursor::with_limits(&good, "test", Limits::new(64, 2));
    assert_eq!(cursor.read_wide_string().unwrap(), "AB");
    cursor.finish().unwrap();

    let mut over_budget = 3_u32.to_le_bytes().to_vec();
    over_budget.extend_from_slice(&[b'A', 0, b'B', 0, b'C', 0]);
    let mut cursor = Cursor::with_limits(&over_budget, "test", Limits::new(64, 2));
    assert!(matches!(
        cursor.read_wide_string(),
        Err(Error::StringLimit {
            units: 3,
            limit: 2,
            ..
        })
    ));

    let mut unpaired = 1_u32.to_le_bytes().to_vec();
    unpaired.extend_from_slice(&0xd800_u16.to_le_bytes());
    let mut cursor = Cursor::new(&unpaired, "test");
    assert!(matches!(
        cursor.read_wide_string(),
        Err(Error::InvalidUtf16 { .. })
    ));
}

#[test]
fn writer_streams_utf16_and_round_trips_without_a_temporary_unit_buffer() {
    let mut bytes = Vec::new();
    Writer::new(&mut bytes)
        .write_wide_string("Hello 世界")
        .unwrap();
    let mut cursor = Cursor::new(&bytes, "test");
    assert_eq!(cursor.read_wide_string().unwrap(), "Hello 世界");
    cursor.finish().unwrap();
}

#[test]
fn byte_blobs_are_lent_from_the_payload() {
    let bytes = [3, 0, 0, 0, 7, 8, 9];
    let mut cursor = Cursor::new(&bytes, "test");
    let blob = cursor.read_blob().unwrap();
    assert_eq!(blob, [7, 8, 9]);
    assert_eq!(blob.as_ptr(), bytes[4..].as_ptr());
}

#[test]
fn trailing_payload_errors_retain_the_callers_context() {
    let cursor = Cursor::new(&[0xaa], "BrtExample");
    assert!(matches!(
        cursor.finish(),
        Err(Error::Trailing {
            context: "BrtExample",
            offset: 0,
            remaining: 1,
        })
    ));
}

#[test]
fn bool32_is_typed_and_rejects_non_boolean_values() {
    let bytes = [1, 0, 0, 0, 2, 0, 0, 0];
    let mut cursor = Cursor::new(&bytes, "test");
    assert!(cursor.read_bool32().unwrap());
    assert!(matches!(
        cursor.read_bool32(),
        Err(Error::InvalidBool {
            value: 2,
            offset: 4
        })
    ));
}

#[test]
fn bool8_is_typed_and_rejects_non_boolean_values() {
    let bytes = [0, 1, 2];
    let mut cursor = Cursor::new(&bytes, "test");
    assert!(!cursor.read_bool8().unwrap());
    assert!(cursor.read_bool8().unwrap());
    assert!(matches!(
        cursor.read_bool8(),
        Err(Error::InvalidBool {
            value: 2,
            offset: 2
        })
    ));
}

fn decode_rk(rk: u32) -> f64 {
    Cursor::new(&rk.to_le_bytes(), "RK").read_rk().unwrap()
}

fn encode_rk(value: f64) -> Result<u32, Error> {
    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write_rk(value)?;
    let encoded = <[u8; 4]>::try_from(bytes.as_slice()).map_err(|_| Error::LengthOverflow {
        what: "RK test output",
        length: bytes.len(),
    })?;
    Ok(u32::from_le_bytes(encoded))
}

#[test]
fn decodes_all_rk_flag_combinations_with_signed_integers() {
    // `[MS-XLSB]` 2.5.123: bit 0 is fx100, bit 1 is fInt, and the
    // remaining bits are either a signed 30-bit integer or the upper 30 f64 bits.
    assert_eq!(decode_rk((42_u32 << 2) | 0x02), 42.0);
    assert_eq!(decode_rk(((-42_i32) as u32) << 2 | 0x02), -42.0);
    assert_eq!(decode_rk((1234_u32 << 2) | 0x03), 12.34);
    assert_eq!(decode_rk(((-1234_i32) as u32) << 2 | 0x03), -12.34);

    let float = ((1.5_f64.to_bits() >> 32) as u32) & 0xffff_fffc;
    assert_eq!(decode_rk(float), 1.5);
    assert_eq!(decode_rk(float | 0x01), 0.015);
}

#[test]
fn writer_uses_only_exact_rk_encodings_at_signed_30_bit_boundaries() {
    let min = f64::from(-(1 << 29));
    let max = f64::from((1 << 29) - 1);
    for value in [min, max, 1.5, 12.34, 1.234375] {
        let rk = encode_rk(value).unwrap();
        assert_eq!(decode_rk(rk).to_bits(), value.to_bits());
    }

    assert_eq!(encode_rk(min).unwrap() & 0x03, 0x02);
    assert_eq!(encode_rk(max).unwrap() & 0x03, 0x02);
    assert_eq!(encode_rk(12.34).unwrap() & 0x03, 0x03);
    assert_eq!(encode_rk(1.234375).unwrap() & 0x03, 0x00);

    for value in [f64::from(i32::MIN), f64::from(1 << 29)] {
        let rk = encode_rk(value).unwrap();
        assert_eq!(rk & 0x02, 0);
        assert_eq!(decode_rk(rk).to_bits(), value.to_bits());
    }
}

#[test]
fn writer_rejects_lossy_rk_values() {
    let value = 1.234_567_89;
    assert!(matches!(
        encode_rk(value),
        Err(Error::UnrepresentableRk { bits }) if bits == value.to_bits()
    ));
}
