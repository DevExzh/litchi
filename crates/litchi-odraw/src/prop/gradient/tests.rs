#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{MAX_STOPS, Position, Stop, encode, parse_payload};
use crate::prop::{ColorRef, Id, Props};
use crate::{Error, Record, RecordKind};

fn opt_record(payload: &[u8], opaque: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(0x8000_u16 | Id::FillShadeColors.raw()).to_le_bytes());
    data.extend_from_slice(
        &i32::try_from(payload.len())
            .expect("payload length fits in i32")
            .to_le_bytes(),
    );
    data.extend_from_slice(&(0x8000_u16 | 0x0600).to_le_bytes());
    data.extend_from_slice(
        &i32::try_from(opaque.len())
            .expect("opaque length fits in i32")
            .to_le_bytes(),
    );
    data.extend_from_slice(payload);
    data.extend_from_slice(opaque);
    data
}

#[test]
fn round_trips_typed_stops_and_retains_neighboring_opaque_properties() {
    let authored = [
        Stop::new(ColorRef::from_raw(0x0033_2211), Position::START),
        Stop::new(
            ColorRef::from_raw(0x0800_0004),
            Position::new(0x0000_8000).expect("halfway"),
        ),
        Stop::new(ColorRef::from_raw(0x00FF_0000), Position::END),
    ];
    let payload = encode(&authored, crate::Limits::default()).expect("encoded stops");
    let parsed = parse_payload(&payload, crate::Limits::default()).expect("parsed stops");

    assert_eq!(parsed.payload(), payload.as_slice());
    assert_eq!(parsed.len(), authored.len());
    assert_eq!(parsed.iter().collect::<Vec<_>>(), authored);

    let opaque = [0xDE, 0xAD, 0xBE, 0xEF];
    let data = opt_record(&payload, &opaque);
    let record = Record::from_parts(RecordKind::Opt, 3, 2, &data, crate::Limits::default())
        .expect("Opt record");
    let properties = Props::parse(&record, crate::Limits::default()).expect("properties");
    let stops = properties
        .gradient_stops()
        .expect("gradient property")
        .expect("stops present");
    assert_eq!(stops.payload(), payload.as_slice());
    assert_eq!(
        properties.get_binary(Id::unknown(0x0600).expect("unknown property")),
        Some(opaque.as_slice())
    );
}

#[test]
fn accepts_empty_and_boundary_positions() {
    let empty = encode(&[], crate::Limits::default()).expect("empty array");
    assert!(
        parse_payload(&empty, crate::Limits::default())
            .expect("empty stops")
            .is_empty()
    );

    let stops = [
        Stop::new(ColorRef::from_raw(0), Position::START),
        Stop::new(ColorRef::from_raw(u32::MAX), Position::END),
    ];
    let payload = encode(&stops, crate::Limits::default()).expect("boundary array");
    let parsed = parse_payload(&payload, crate::Limits::default()).expect("stops");
    assert_eq!(parsed.get(0), Some(stops[0]));
    assert_eq!(parsed.get(1), Some(stops[1]));
}

#[test]
fn rejects_bad_width_range_order_and_amplification() {
    let mut bad_width = Vec::new();
    bad_width.extend_from_slice(&1_u16.to_le_bytes());
    bad_width.extend_from_slice(&1_u16.to_le_bytes());
    bad_width.extend_from_slice(&4_u16.to_le_bytes());
    bad_width.extend_from_slice(&[0; 4]);
    assert!(matches!(
        parse_payload(&bad_width, crate::Limits::default()),
        Err(Error::MalformedProperties {
            reason: "gradient stop array element size is not MSOSHADECOLOR"
        })
    ));

    let mut out_of_range = encode(
        &[Stop::new(ColorRef::from_raw(1), Position::START)],
        crate::Limits::default(),
    )
    .expect("valid base");
    out_of_range[10..14].copy_from_slice(&(-1_i32).to_le_bytes());
    assert!(matches!(
        parse_payload(&out_of_range, crate::Limits::default()),
        Err(Error::MalformedProperties {
            reason: "gradient stop position is outside the inclusive 0..1 range"
        })
    ));

    let mut duplicate = encode(
        &[
            Stop::new(ColorRef::from_raw(1), Position::START),
            Stop::new(
                ColorRef::from_raw(2),
                Position::new(1).expect("one fixed-point unit"),
            ),
        ],
        crate::Limits::default(),
    )
    .expect("valid base");
    duplicate[18..22].copy_from_slice(&0_i32.to_le_bytes());
    assert!(matches!(
        parse_payload(&duplicate, crate::Limits::default()),
        Err(Error::MalformedProperties {
            reason: "gradient stop positions are not strictly ascending"
        })
    ));

    let count = u16::try_from(MAX_STOPS)
        .expect("bound fits in u16")
        .saturating_add(1);
    let mut too_many = Vec::new();
    too_many.extend_from_slice(&count.to_le_bytes());
    too_many.extend_from_slice(&count.to_le_bytes());
    too_many.extend_from_slice(&8_u16.to_le_bytes());
    too_many.resize(6 + usize::from(count) * 8, 0);
    assert!(matches!(
        parse_payload(&too_many, crate::Limits::default()),
        Err(Error::MalformedProperties {
            reason: "gradient stop count exceeds the safe bound"
        })
    ));

    let mut simple = Vec::new();
    simple.extend_from_slice(&Id::FillShadeColors.raw().to_le_bytes());
    simple.extend_from_slice(&1_i32.to_le_bytes());
    let record = Record::from_parts(RecordKind::Opt, 3, 1, &simple, crate::Limits::default())
        .expect("Opt record");
    let properties = Props::parse(&record, crate::Limits::default()).expect("property table");
    assert!(matches!(
        properties.gradient_stops(),
        Err(Error::MalformedProperties {
            reason: "fillShadeColors must be a complex IMsoArray property"
        })
    ));
}
