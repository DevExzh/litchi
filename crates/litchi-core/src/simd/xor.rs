//! SIMD-accelerated XOR operations
//!
//! Provides high-performance XOR operations using SIMD instructions when available.
//! Automatically selects the best implementation based on CPU features.

#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::{
    _mm_loadu_si128, _mm_storeu_si128, _mm_xor_si128, _mm256_broadcastsi128_si256,
    _mm256_loadu_si256, _mm256_storeu_si256, _mm256_xor_si256,
};

#[cfg(target_arch = "aarch64")]
use std::arch::aarch64::*;

/// XOR two 16-byte arrays and store the result in the destination.
///
/// This function uses SIMD instructions when available:
/// - `x86_64`: SSE2 (128-bit)
/// - aarch64: NEON (128-bit, always available)
/// - Other: Scalar fallback
///
/// # Panics
///
/// Panics unless `dst` and `src` each contain exactly 16 bytes.
#[inline]
#[allow(
    clippy::module_name_repetitions,
    reason = "public API name is stable and used by dependent crates; renaming would be a breaking change"
)]
pub fn xor_16_bytes(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    assert_eq!(dst.len(), 16, "destination must contain exactly 16 bytes");
    assert_eq!(src.len(), 16, "source must contain exactly 16 bytes");

    #[cfg(target_arch = "x86_64")]
    {
        if is_x86_feature_detected!("sse2") {
            // SAFETY: SSE2 support was detected at runtime, and the assertions
            // above guarantee that `dst` and `src` are 16 bytes long.
            unsafe { xor_16_bytes_sse2(dst, src, key) }
        } else {
            xor_16_bytes_scalar(dst, src, key);
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
        // SAFETY: NEON is part of the aarch64 baseline, and the assertions
        // above guarantee that both raw vector accesses are in bounds.
        unsafe { xor_16_bytes_neon(dst, src, key) }
    }

    #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
    {
        xor_16_bytes_scalar(dst, src, key);
    }
}

/// XOR 32 bytes in place with a 16-byte key (repeated twice).
///
/// This function uses SIMD instructions when available:
/// - `x86_64`: AVX2 (256-bit, single operation) or SSE2 (128-bit, two operations)
/// - aarch64: NEON (128-bit, two operations)
/// - Other: Scalar fallback
///
/// # Panics
///
/// Panics unless `data` contains exactly 32 bytes.
#[inline]
#[allow(
    clippy::module_name_repetitions,
    reason = "public API name is stable and used by dependent crates; renaming would be a breaking change"
)]
pub fn xor_32_bytes_inplace(data: &mut [u8], key: &[u8; 16]) {
    assert_eq!(data.len(), 32, "data must contain exactly 32 bytes");

    #[cfg(target_arch = "x86_64")]
    {
        if is_x86_feature_detected!("avx2") {
            // SAFETY: AVX2 support was detected at runtime, and the assertion
            // above guarantees that `data` is 32 bytes long.
            unsafe { xor_32_bytes_inplace_avx2(data, key) }
        } else if is_x86_feature_detected!("sse2") {
            // SAFETY: SSE2 support was detected at runtime, and the assertion
            // above guarantees that `data` is 32 bytes long.
            unsafe { xor_32_bytes_inplace_sse2(data, key) }
        } else {
            xor_32_bytes_inplace_scalar(data, key);
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
        // SAFETY: NEON is part of the aarch64 baseline, and the assertion above
        // guarantees that both 16-byte vector accesses are in bounds.
        unsafe { xor_32_bytes_inplace_neon(data, key) }
    }

    #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
    {
        xor_32_bytes_inplace_scalar(data, key);
    }
}

// === x86_64 implementations ===

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "sse2")]
#[inline]
unsafe fn xor_16_bytes_sse2(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    // SAFETY: the caller guarantees `src` is at least 16 readable bytes;
    // `loadu` does not require alignment.
    let src_vec = unsafe { _mm_loadu_si128(src.as_ptr().cast()) };
    // SAFETY: `key` is a 16-byte array; `loadu` does not require alignment.
    let key_vec = unsafe { _mm_loadu_si128(key.as_ptr().cast()) };
    let result = _mm_xor_si128(src_vec, key_vec);
    // SAFETY: the caller guarantees `dst` is at least 16 writable bytes;
    // `storeu` does not require alignment.
    unsafe { _mm_storeu_si128(dst.as_mut_ptr().cast(), result) };
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "sse2")]
#[inline]
unsafe fn xor_32_bytes_inplace_sse2(data: &mut [u8], key: &[u8; 16]) {
    // SAFETY: `key` is a 16-byte array; `loadu` does not require alignment.
    let key_vec = unsafe { _mm_loadu_si128(key.as_ptr().cast()) };

    // XOR first 16 bytes
    // SAFETY: the caller guarantees `data` is at least 32 bytes; `loadu` does
    // not require alignment.
    let data_vec1 = unsafe { _mm_loadu_si128(data.as_ptr().cast()) };
    let result1 = _mm_xor_si128(data_vec1, key_vec);
    // SAFETY: same bounds argument as above; `storeu` does not require alignment.
    unsafe { _mm_storeu_si128(data.as_mut_ptr().cast(), result1) };

    // XOR second 16 bytes
    // SAFETY: the second 16-byte half of `data` is readable; `loadu` does not
    // require alignment.
    let data_vec2 = unsafe { _mm_loadu_si128(data[16..].as_ptr().cast()) };
    let result2 = _mm_xor_si128(data_vec2, key_vec);
    // SAFETY: the second 16-byte half of `data` is writable; `storeu` does not
    // require alignment.
    unsafe { _mm_storeu_si128(data[16..].as_mut_ptr().cast(), result2) };
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "avx2")]
#[inline]
unsafe fn xor_32_bytes_inplace_avx2(data: &mut [u8], key: &[u8; 16]) {
    // Broadcast the 16-byte key to 32 bytes (duplicate it)
    // SAFETY: `key` is a 16-byte array; `loadu` does not require alignment.
    let key_low = unsafe { _mm_loadu_si128(key.as_ptr().cast()) };
    let key_256 = _mm256_broadcastsi128_si256(key_low);

    // Load, XOR, and store 32 bytes in one go
    // SAFETY: the caller guarantees `data` is at least 32 bytes; `loadu` does
    // not require alignment.
    let data_vec = unsafe { _mm256_loadu_si256(data.as_ptr().cast()) };
    let result = _mm256_xor_si256(data_vec, key_256);
    // SAFETY: same bounds argument as above; `storeu` does not require alignment.
    unsafe { _mm256_storeu_si256(data.as_mut_ptr().cast(), result) };
}

// === aarch64 implementations ===

#[cfg(target_arch = "aarch64")]
#[inline]
unsafe fn xor_16_bytes_neon(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    unsafe {
        let src_vec = vld1q_u8(src.as_ptr());
        let key_vec = vld1q_u8(key.as_ptr());
        let result = veorq_u8(src_vec, key_vec);
        vst1q_u8(dst.as_mut_ptr(), result);
    }
}

#[cfg(target_arch = "aarch64")]
#[inline]
unsafe fn xor_32_bytes_inplace_neon(data: &mut [u8], key: &[u8; 16]) {
    unsafe {
        let key_vec = vld1q_u8(key.as_ptr());

        // XOR first 16 bytes
        let data_vec1 = vld1q_u8(data.as_ptr());
        let result1 = veorq_u8(data_vec1, key_vec);
        vst1q_u8(data.as_mut_ptr(), result1);

        // XOR second 16 bytes
        let data_vec2 = vld1q_u8(data.as_ptr().add(16));
        let result2 = veorq_u8(data_vec2, key_vec);
        vst1q_u8(data.as_mut_ptr().add(16), result2);
    }
}

// === Scalar fallback implementations ===

#[cfg(any(test, not(target_arch = "aarch64")))]
#[inline]
fn xor_16_bytes_scalar(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    for i in 0..16 {
        dst[i] = src[i] ^ key[i];
    }
}

#[cfg(any(test, not(target_arch = "aarch64")))]
#[inline]
fn xor_32_bytes_inplace_scalar(data: &mut [u8], key: &[u8; 16]) {
    for i in 0..16 {
        data[i] ^= key[i];
        data[i + 16] ^= key[i];
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_xor_16_bytes() {
        let src = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10,
        ];
        let key = [
            0xFF, 0xFE, 0xFD, 0xFC, 0xFB, 0xFA, 0xF9, 0xF8, 0xF7, 0xF6, 0xF5, 0xF4, 0xF3, 0xF2,
            0xF1, 0xF0,
        ];
        let mut dst = [0u8; 16];
        let mut scalar_dst = [0u8; 16];

        xor_16_bytes(&mut dst, &src, &key);
        xor_16_bytes_scalar(&mut scalar_dst, &src, &key);

        assert_eq!(dst, scalar_dst);
    }

    #[test]
    #[should_panic(expected = "destination must contain exactly 16 bytes")]
    fn test_xor_16_bytes_rejects_short_destination() {
        xor_16_bytes(&mut [0u8; 15], &[0u8; 16], &[0u8; 16]);
    }

    #[test]
    #[should_panic(expected = "source must contain exactly 16 bytes")]
    fn test_xor_16_bytes_rejects_short_source() {
        xor_16_bytes(&mut [0u8; 16], &[0u8; 15], &[0u8; 16]);
    }

    #[test]
    fn test_xor_32_bytes_inplace() {
        let mut data = [
            0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B, 0x0C, 0x0D, 0x0E,
            0x0F, 0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x1B, 0x1C,
            0x1D, 0x1E, 0x1F, 0x20,
        ];
        let key = [
            0xFF, 0xFE, 0xFD, 0xFC, 0xFB, 0xFA, 0xF9, 0xF8, 0xF7, 0xF6, 0xF5, 0xF4, 0xF3, 0xF2,
            0xF1, 0xF0,
        ];
        let original = data;
        let mut scalar_data = original;

        xor_32_bytes_inplace(&mut data, &key);
        xor_32_bytes_inplace_scalar(&mut scalar_data, &key);

        assert_eq!(data, scalar_data);

        // XOR again should restore original (XOR is reversible)
        xor_32_bytes_inplace(&mut data, &key);
        assert_eq!(data, original);
    }

    #[test]
    #[should_panic(expected = "data must contain exactly 32 bytes")]
    fn test_xor_32_bytes_inplace_rejects_short_input() {
        xor_32_bytes_inplace(&mut [0u8; 31], &[0u8; 16]);
    }

    #[test]
    fn test_xor_reversibility() {
        let mut data = [0xAAu8; 32];
        let key = [0x55u8; 16];
        let original = data;

        xor_32_bytes_inplace(&mut data, &key);
        assert_ne!(data, original);

        xor_32_bytes_inplace(&mut data, &key);
        assert_eq!(data, original);
    }
}
