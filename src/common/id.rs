use crate::common::simd::fmt::hex_encode_to_string;
use rand::Rng;

/// Generate a random RFC4122 v4 GUID as raw 16 bytes
pub fn generate_guid_bytes() -> [u8; 16] {
    let mut bytes = [0u8; 16];
    let mut rng = rand::rng();
    rng.fill(&mut bytes);
    // RFC4122 v4
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    bytes
}

/// Generate a random GUID in the form {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
pub fn generate_guid_braced() -> String {
    let bytes = generate_guid_bytes();
    format_guid_braced(&bytes)
}

/// Format raw GUID bytes as a braced string {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
pub fn format_guid_braced(bytes: &[u8; 16]) -> String {
    let mut out = String::with_capacity(38);
    out.push('{');
    hex_encode_to_string(&bytes[0..4], &mut out, false);
    out.push('-');
    hex_encode_to_string(&bytes[4..6], &mut out, false);
    out.push('-');
    hex_encode_to_string(&bytes[6..8], &mut out, false);
    out.push('-');
    hex_encode_to_string(&bytes[8..10], &mut out, false);
    out.push('-');
    hex_encode_to_string(&bytes[10..16], &mut out, false);
    out.push('}');
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_generate_guid_braced_format() {
        let s = generate_guid_braced();
        assert_eq!(s.len(), 38);
        assert!(s.starts_with('{'));
        assert!(s.ends_with('}'));
        assert_eq!(&s[9..10], "-");
        assert_eq!(&s[14..15], "-");
        assert_eq!(&s[19..20], "-");
        assert_eq!(&s[24..25], "-");
        for (i, ch) in s.chars().enumerate() {
            if matches!(i, 0 | 37 | 9 | 14 | 19 | 24) {
                continue;
            }
            assert!(ch.is_ascii_hexdigit());
            if ch.is_ascii_alphabetic() {
                assert!(ch.is_ascii_uppercase());
            }
        }
    }
}
