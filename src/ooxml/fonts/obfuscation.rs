use crate::common::encoding::decode_hex_data;
use crate::common::simd::xor::xor_32_bytes_inplace;
use crate::ooxml::error::{OoxmlError, Result};

/// Obfuscates font data according to OOXML specification (ISO/IEC 29500-1:2016, 15.2.14).
///
/// The obfuscation is a simple XOR of the first 32 bytes of the font data
/// with the GUID bytes, processed in reversed byte order.
///
/// This implementation uses SIMD instructions (AVX2/SSE2/NEON) when available
/// for optimal performance, with automatic fallback to scalar code.
#[inline]
pub fn obfuscate_font_data_bytes(data: &mut [u8], guid_bytes: &[u8; 16]) {
    if data.len() < 32 {
        return;
    }

    // The key is derived from the GUID by reversing its byte order
    let key = [
        guid_bytes[15],
        guid_bytes[14],
        guid_bytes[13],
        guid_bytes[12],
        guid_bytes[11],
        guid_bytes[10],
        guid_bytes[9],
        guid_bytes[8],
        guid_bytes[7],
        guid_bytes[6],
        guid_bytes[5],
        guid_bytes[4],
        guid_bytes[3],
        guid_bytes[2],
        guid_bytes[1],
        guid_bytes[0],
    ];

    // XOR the first 32 bytes with the 16-byte key (repeated twice)
    // Uses SIMD acceleration (AVX2/SSE2 on x86_64, NEON on aarch64)
    xor_32_bytes_inplace(&mut data[..32], &key);
}

/// De-obfuscates font data. Since it's XOR, it's the same operation as obfuscation.
#[inline]
pub fn deobfuscate_font_data_bytes(data: &mut [u8], guid_bytes: &[u8; 16]) {
    obfuscate_font_data_bytes(data, guid_bytes)
}

/// Obfuscates font data using a GUID string (for backward compatibility).
///
/// For better performance, prefer using `obfuscate_font_data_bytes` directly.
pub fn obfuscate_font_data(data: &mut [u8], guid_str: &str) -> Result<()> {
    if data.len() < 32 {
        return Ok(());
    }

    // Parse GUID string: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} or XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
    let clean_guid = guid_str
        .trim_matches(|c| c == '{' || c == '}')
        .replace('-', "");
    let guid_bytes = decode_hex_data(&clean_guid)
        .map_err(|e| OoxmlError::Other(format!("Invalid GUID format: {}", e)))?;

    if guid_bytes.len() != 16 {
        return Err(OoxmlError::Other(format!(
            "Invalid GUID length: expected 16 bytes, got {}",
            guid_bytes.len()
        )));
    }

    let guid_array: [u8; 16] = guid_bytes.try_into().unwrap();
    obfuscate_font_data_bytes(data, &guid_array);
    Ok(())
}

/// De-obfuscates font data using a GUID string.
pub fn deobfuscate_font_data(data: &mut [u8], guid_str: &str) -> Result<()> {
    obfuscate_font_data(data, guid_str)
}
