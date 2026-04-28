//! SIMD-accelerated XOR operations
//!
//! Provides high-performance XOR operations using SIMD instructions when available.
//! Automatically selects the best implementation based on CPU features.

#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::*;

#[cfg(target_arch = "aarch64")]
use std::arch::aarch64::*;

/// XOR two 16-byte arrays and store the result in the destination.
///
/// This function uses SIMD instructions when available:
/// - x86_64: SSE2 (128-bit)
/// - aarch64: NEON (128-bit, always available)
/// - Other: Scalar fallback
#[inline]
pub fn xor_16_bytes(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    debug_assert_eq!(dst.len(), 16);
    debug_assert_eq!(src.len(), 16);

    #[cfg(target_arch = "x86_64")]
    {
        if is_x86_feature_detected!("sse2") {
            unsafe { xor_16_bytes_sse2(dst, src, key) }
        } else {
            xor_16_bytes_scalar(dst, src, key);
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
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
/// - x86_64: AVX2 (256-bit, single operation) or SSE2 (128-bit, two operations)
/// - aarch64: NEON (128-bit, two operations)
/// - Other: Scalar fallback
#[inline]
pub fn xor_32_bytes_inplace(data: &mut [u8], key: &[u8; 16]) {
    debug_assert_eq!(data.len(), 32);

    #[cfg(target_arch = "x86_64")]
    {
        if is_x86_feature_detected!("avx2") {
            unsafe { xor_32_bytes_inplace_avx2(data, key) }
        } else if is_x86_feature_detected!("sse2") {
            unsafe { xor_32_bytes_inplace_sse2(data, key) }
        } else {
            xor_32_bytes_inplace_scalar(data, key);
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
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
    unsafe {
        let src_vec = _mm_loadu_si128(src.as_ptr() as *const __m128i);
        let key_vec = _mm_loadu_si128(key.as_ptr() as *const __m128i);
        let result = _mm_xor_si128(src_vec, key_vec);
        _mm_storeu_si128(dst.as_mut_ptr() as *mut __m128i, result);
    }
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "sse2")]
#[inline]
unsafe fn xor_32_bytes_inplace_sse2(data: &mut [u8], key: &[u8; 16]) {
    unsafe {
        let key_vec = _mm_loadu_si128(key.as_ptr() as *const __m128i);

        // XOR first 16 bytes
        let data_vec1 = _mm_loadu_si128(data.as_ptr() as *const __m128i);
        let result1 = _mm_xor_si128(data_vec1, key_vec);
        _mm_storeu_si128(data.as_mut_ptr() as *mut __m128i, result1);

        // XOR second 16 bytes
        let data_vec2 = _mm_loadu_si128(data.as_ptr().add(16) as *const __m128i);
        let result2 = _mm_xor_si128(data_vec2, key_vec);
        _mm_storeu_si128(data.as_mut_ptr().add(16) as *mut __m128i, result2);
    }
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "avx2")]
#[inline]
unsafe fn xor_32_bytes_inplace_avx2(data: &mut [u8], key: &[u8; 16]) {
    unsafe {
        // Broadcast the 16-byte key to 32 bytes (duplicate it)
        let key_low = _mm_loadu_si128(key.as_ptr() as *const __m128i);
        let key_256 = _mm256_broadcastsi128_si256(key_low);

        // Load, XOR, and store 32 bytes in one go
        let data_vec = _mm256_loadu_si256(data.as_ptr() as *const __m256i);
        let result = _mm256_xor_si256(data_vec, key_256);
        _mm256_storeu_si256(data.as_mut_ptr() as *mut __m256i, result);
    }
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

#[inline]
fn xor_16_bytes_scalar(dst: &mut [u8], src: &[u8], key: &[u8; 16]) {
    for i in 0..16 {
        dst[i] = src[i] ^ key[i];
    }
}

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

        xor_16_bytes(&mut dst, &src, &key);

        // Verify XOR result
        for i in 0..16 {
            assert_eq!(dst[i], src[i] ^ key[i]);
        }
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

        xor_32_bytes_inplace(&mut data, &key);

        // Verify XOR result (key is repeated twice)
        for i in 0..16 {
            assert_eq!(data[i], original[i] ^ key[i]);
            assert_eq!(data[i + 16], original[i + 16] ^ key[i]);
        }

        // XOR again should restore original (XOR is reversible)
        xor_32_bytes_inplace(&mut data, &key);
        assert_eq!(data, original);
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
