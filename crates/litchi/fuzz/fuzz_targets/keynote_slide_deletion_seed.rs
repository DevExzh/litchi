//! Reproducible native three-slide A/B/C source for the deletion target.
//!
//! The payload is checked into the fuzz target as text so the target does not
//! depend on a developer's private temporary directory or on integration-test
//! builders. Decoding happens once before the immutable package is shared by
//! all fuzz iterations.

use std::sync::OnceLock;

const SEED_BASE64: &str = include_str!("keynote_slide_deletion_seed.b64");

pub fn bytes() -> &'static [u8] {
    static SEED: OnceLock<Box<[u8]>> = OnceLock::new();
    SEED.get_or_init(|| decode_base64(SEED_BASE64)).as_ref()
}

fn decode_base64(input: &str) -> Box<[u8]> {
    let mut output = Vec::with_capacity(input.len().saturating_mul(3) / 4);
    let mut accumulator = 0_u32;
    let mut bits = 0_u8;
    let mut padding = false;

    for byte in input.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        if byte == b'=' {
            padding = true;
            continue;
        }
        assert!(!padding, "invalid base64 seed padding");
        let value = match byte {
            b'A'..=b'Z' => byte - b'A',
            b'a'..=b'z' => byte - b'a' + 26,
            b'0'..=b'9' => byte - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            _ => panic!("invalid base64 seed byte"),
        };
        accumulator = (accumulator << 6) | u32::from(value);
        bits = bits.saturating_add(6);
        if bits >= 8 {
            bits -= 8;
            output.push((accumulator >> bits) as u8);
            accumulator &= (1_u32 << bits).saturating_sub(1);
        }
    }

    assert!(
        !output.is_empty(),
        "Keynote deletion seed must not be empty"
    );
    output.into_boxed_slice()
}
