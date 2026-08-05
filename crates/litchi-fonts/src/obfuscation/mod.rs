use litchi_core::hex;
use litchi_core::simd::xor::xor_32_bytes_inplace;

/// Errors produced while parsing a font obfuscation key.
#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum Error {
    #[error("invalid GUID format: {0}")]
    InvalidGuid(String),
    #[error("invalid GUID length: expected 16 bytes, got {0}")]
    InvalidGuidLength(usize),
}

/// Obfuscate font data according to OOXML (ISO/IEC 29500-1:2016, 15.2.14).
///
/// The obfuscation is a simple XOR of the first 32 bytes of the font data
/// with the GUID bytes, processed in reversed byte order.
///
/// This implementation uses SIMD instructions (AVX2/SSE2/NEON) when available
/// for optimal performance, with automatic fallback to scalar code.
#[inline]
pub fn apply_bytes(data: &mut [u8], key: &[u8; 16]) {
    if data.len() < 32 {
        return;
    }

    // The key is derived from the GUID by reversing its byte order
    let key = [
        key[15], key[14], key[13], key[12], key[11], key[10], key[9], key[8], key[7], key[6],
        key[5], key[4], key[3], key[2], key[1], key[0],
    ];

    // XOR the first 32 bytes with the 16-byte key (repeated twice)
    // Uses SIMD acceleration (AVX2/SSE2 on x86_64, NEON on aarch64)
    xor_32_bytes_inplace(&mut data[..32], &key);
}

/// De-obfuscates font data. Since it's XOR, it's the same operation as obfuscation.
#[inline]
pub fn remove_bytes(data: &mut [u8], key: &[u8; 16]) {
    apply_bytes(data, key)
}

/// Obfuscate font data using a GUID string.
pub fn apply(data: &mut [u8], guid: &str) -> Result<(), Error> {
    if data.len() < 32 {
        return Ok(());
    }

    // Parse GUID string: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} or XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
    let clean_guid = guid.trim_matches(|c| c == '{' || c == '}').replace('-', "");
    let guid_bytes =
        hex::decode(&clean_guid).map_err(|error| Error::InvalidGuid(error.to_string()))?;

    let guid_array: [u8; 16] = guid_bytes
        .try_into()
        .map_err(|bytes: Vec<u8>| Error::InvalidGuidLength(bytes.len()))?;
    apply_bytes(data, &guid_array);
    Ok(())
}

/// Remove the OOXML obfuscation using a GUID string.
pub fn remove(data: &mut [u8], guid: &str) -> Result<(), Error> {
    apply(data, guid)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn guid_api_round_trips_and_rejects_invalid_input() {
        let original = [0x5a; 32];
        let mut data = original;
        let guid = "{00112233-4455-6677-8899-aabbccddeeff}";

        apply(&mut data, guid).expect("valid GUID");
        assert_ne!(data, original);
        remove(&mut data, guid).expect("valid GUID");
        assert_eq!(data, original);

        assert!(apply(&mut data, "not-a-guid").is_err());
        assert_eq!(apply(&mut data, "0011"), Err(Error::InvalidGuidLength(2)));
    }
}
