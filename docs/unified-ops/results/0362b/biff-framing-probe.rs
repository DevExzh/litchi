//! Focused raw framing probe for record 0362b; see biff-validation.json.
use litchi_biff::{Encoder, Kind, Limits, Records};
use std::hint::black_box;
use std::time::Instant;

fn scan_default(bytes: &[u8]) -> u64 {
    let records = Records::new(black_box(bytes), Limits::default()).expect("valid limits");
    checksum(records)
}

fn scan_configured(bytes: &[u8]) -> u64 {
    let records =
        Records::new(black_box(bytes), black_box(Limits::default())).expect("valid limits");
    checksum(records)
}

fn checksum(records: Records<'_>) -> u64 {
    records.fold(0_u64, |sum, record| {
        let record = record.expect("valid generated frame");
        sum.wrapping_add(record.offset() as u64)
            .wrapping_add(u64::from(record.kind().get()))
            .wrapping_add(u64::from(record.payload()[0]))
    })
}

fn encode_default() -> u64 {
    let encoder = Encoder::new(Limits::default()).expect("valid limits");
    encode(encoder)
}

fn encode_configured() -> u64 {
    let encoder = Encoder::new(black_box(Limits::default())).expect("valid limits");
    encode(encoder)
}

fn encode(mut encoder: Encoder) -> u64 {
    for kind in 0..128_u16 {
        encoder
            .push(Kind::from_wire(kind), black_box(&[0xa5; 64]))
            .expect("bounded frame");
    }
    let bytes = black_box(encoder.finish());
    bytes.len() as u64 + u64::from(bytes[2]) + u64::from(bytes[8703])
}

fn main() {
    let mut args = std::env::args().skip(1);
    let route = args.next().expect("route");
    let iterations: usize = args.next().expect("iterations").parse().expect("integer");
    let mut input = Vec::new();
    for kind in 0..32_768_u16 {
        input.extend_from_slice(&kind.to_le_bytes());
        input.extend_from_slice(&16_u16.to_le_bytes());
        input.extend_from_slice(&[0xa5; 16]);
    }
    let operation = || match route.as_str() {
        "scan-default" => scan_default(&input),
        "scan-configured" => scan_configured(&input),
        "encode-default" => encode_default(),
        "encode-configured" => encode_configured(),
        _ => panic!("unknown route"),
    };
    let expected = operation();
    for _ in 0..8 {
        assert_eq!(black_box(operation()), expected);
    }
    let started = Instant::now();
    let mut sum = 0_u64;
    for _ in 0..iterations {
        sum = sum.wrapping_add(black_box(operation()));
    }
    let elapsed = started.elapsed().as_nanos();
    assert_eq!(sum, expected.wrapping_mul(iterations as u64));
    println!(
        "{{\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{elapsed},\"checksum\":{sum},\"records_size\":{}}}",
        std::mem::size_of::<Records<'_>>()
    );
}
