use super::codec::{
    CRT_LAYOUT_12_A_RECORD_TYPE, CRT_LAYOUT_12_RECORD_TYPE, FRT_HEADER_LEN, LayoutReader,
};
use super::model::{LayoutModes, XlsCrtLayout12, XlsCrtLayout12A, XlsCrtLayout12Mode};

fn layout12_record(checksum: u32, flags: u16, modes: [u16; 4], values: [f64; 4]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&CRT_LAYOUT_12_RECORD_TYPE.to_le_bytes());
    data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
    data.extend_from_slice(&checksum.to_le_bytes());
    data.extend_from_slice(&flags.to_le_bytes());
    for mode in modes {
        data.extend_from_slice(&mode.to_le_bytes());
    }
    for value in values {
        data.extend_from_slice(&value.to_le_bytes());
    }
    data.extend_from_slice(&[0; 2]);
    data
}

fn layout12a_record(
    checksum: u32,
    flags: u16,
    corners: [i16; 4],
    modes: [u16; 4],
    values: [f64; 4],
) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&CRT_LAYOUT_12_A_RECORD_TYPE.to_le_bytes());
    data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
    data.extend_from_slice(&checksum.to_le_bytes());
    data.extend_from_slice(&flags.to_le_bytes());
    for corner in corners {
        data.extend_from_slice(&corner.to_le_bytes());
    }
    for mode in modes {
        data.extend_from_slice(&mode.to_le_bytes());
    }
    for value in values {
        data.extend_from_slice(&value.to_le_bytes());
    }
    data.extend_from_slice(&[0; 2]);
    data
}

#[test]
fn layout_modes_rejects_truncated_and_overflowing_reads() {
    let mut truncated = LayoutReader::new(&[0; 39], CRT_LAYOUT_12_RECORD_TYPE);
    assert!(LayoutModes::parse(&mut truncated).is_err());
    assert_eq!(truncated.offset, 32);

    let mut overflowing = LayoutReader {
        data: &[],
        offset: usize::MAX,
        record_type: CRT_LAYOUT_12_RECORD_TYPE,
    };
    assert!(LayoutModes::parse(&mut overflowing).is_err());
    assert_eq!(overflowing.offset, usize::MAX);
}

#[test]
fn crt_layout12_round_trip() {
    let bytes = layout12_record(
        0x0000_4321,
        0x0008,
        [0x0001, 0x0001, 0x0002, 0x0002],
        [0.25, -0.5, 0.75, 1.0],
    );
    let parsed = XlsCrtLayout12::parse(&bytes).unwrap();
    assert_eq!(parsed.checksum(), 0x0000_4321);
    assert_eq!(parsed.auto_layout_type(), 0x4);
    assert_eq!(parsed.x_mode(), XlsCrtLayout12Mode::Factor);
    assert_eq!(parsed.width_mode(), XlsCrtLayout12Mode::Edge);
    assert_eq!(parsed.x(), 0.25);
    assert_eq!(parsed.y(), -0.5);
    assert_eq!(parsed.dx(), 0.75);
    assert_eq!(parsed.dy(), 1.0);
    assert_eq!(parsed.to_payload(), bytes);
}

#[test]
fn crt_layout12_preserves_unused_and_reserved_bits() {
    // The unused bit, the 11 reserved1 bits, and reserved2 MUST be ignored
    // but round-trip verbatim.
    let mut bytes = layout12_record(0, 0xF801, [0; 4], [0.0; 4]);
    bytes[58..60].copy_from_slice(&0x7F7Fu16.to_le_bytes());
    bytes[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
    let parsed = XlsCrtLayout12::parse(&bytes).unwrap();
    assert_eq!(parsed.flags(), 0xF801);
    assert_eq!(parsed.auto_layout_type(), 0);
    assert_eq!(parsed.to_payload(), bytes);
}

#[test]
fn crt_layout12_rejects_malformed_records() {
    let bytes = layout12_record(0, 0, [0; 4], [0.0; 4]);
    // Truncated and overlong payloads.
    assert!(XlsCrtLayout12::parse(&bytes[..59]).is_err());
    assert!(XlsCrtLayout12::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
    // Wrong FrtHeader.rt.
    let mut wrong_rt = bytes.clone();
    wrong_rt[0..2].copy_from_slice(&0x089Eu16.to_le_bytes());
    assert!(XlsCrtLayout12::parse(&wrong_rt).is_err());
    // fFrtRef / fFrtAlert set.
    let mut bad_flags = bytes.clone();
    bad_flags[2..4].copy_from_slice(&0x0001u16.to_le_bytes());
    assert!(XlsCrtLayout12::parse(&bad_flags).is_err());
    // Undefined layout mode.
    assert!(XlsCrtLayout12::parse(&layout12_record(0, 0, [3, 0, 0, 0], [0.0; 4])).is_err());
}

#[test]
fn crt_layout12a_round_trip() {
    for checksum in [0, 1] {
        let bytes = layout12a_record(
            checksum,
            0x0001,
            [100, -50, 4000, 3000],
            [0x0000, 0x0001, 0x0002, 0x0000],
            [0.1, 0.2, 0.3, 0.4],
        );
        let parsed = XlsCrtLayout12A::parse(&bytes).unwrap();
        assert_eq!(parsed.checksum(), checksum);
        assert!(parsed.is_layout_target_inner());
        assert_eq!(parsed.x_top_left(), 100);
        assert_eq!(parsed.y_top_left(), -50);
        assert_eq!(parsed.x_bottom_right(), 4000);
        assert_eq!(parsed.y_bottom_right(), 3000);
        assert_eq!(parsed.y_mode(), XlsCrtLayout12Mode::Factor);
        assert_eq!(parsed.width_mode(), XlsCrtLayout12Mode::Edge);
        assert_eq!(parsed.dx(), 0.3);
        assert_eq!(parsed.to_payload(), bytes);
    }
}

#[test]
fn crt_layout12a_rejects_malformed_records() {
    let bytes = layout12a_record(0, 0, [0; 4], [0; 4], [0.0; 4]);
    // Truncated.
    assert!(XlsCrtLayout12A::parse(&bytes[..67]).is_err());
    // Wrong FrtHeader.rt.
    let mut wrong_rt = bytes.clone();
    wrong_rt[0..2].copy_from_slice(&0x089Du16.to_le_bytes());
    assert!(XlsCrtLayout12A::parse(&wrong_rt).is_err());
    // dwCheckSum outside 0x00000000..=0x00000001.
    assert!(XlsCrtLayout12A::parse(&layout12a_record(2, 0, [0; 4], [0; 4], [0.0; 4])).is_err());
    assert!(
        XlsCrtLayout12A::parse(&layout12a_record(0xFFFF_FFFF, 0, [0; 4], [0; 4], [0.0; 4]))
            .is_err()
    );
    // Undefined layout mode.
    assert!(
        XlsCrtLayout12A::parse(&layout12a_record(0, 0, [0; 4], [0, 9, 0, 0], [0.0; 4])).is_err()
    );
}
