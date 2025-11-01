//! SIMD vector comparison operations
//!
//! This module provides high-performance SIMD comparison operations for various
//! architectures (x86_64, aarch64) and instruction sets (AVX, AVX2, AVX512, SSE, NEON, SVE).
//!
//! # Examples
//!
//! ```rust
//! use litchi::common::simd::cmp::{simd_eq_u8, simd_ne_u8};
//!
//! let a = vec![1u8, 2, 3, 4, 5, 6, 7, 8];
//! let b = vec![1u8, 2, 0, 4, 5, 0, 7, 8];
//! let mut result = vec![0u8; 8];
//!
//! simd_eq_u8(&a, &b, &mut result);
//! // result contains: [255, 255, 0, 255, 255, 0, 255, 255]
//! ```

#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::*;

#[cfg(target_arch = "aarch64")]
use std::arch::aarch64::*;

/// Result of a SIMD comparison operation
///
/// The mask is represented as a vector where each element is either all 1s (true)
/// or all 0s (false).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SimdMask {
    /// The underlying mask bits
    pub mask: u64,
}

impl SimdMask {
    /// Create a new SIMD mask from raw bits
    #[inline]
    pub const fn new(mask: u64) -> Self {
        Self { mask }
    }

    /// Check if all bits are set
    #[inline]
    pub const fn all(self) -> bool {
        self.mask == u64::MAX
    }

    /// Check if any bit is set
    #[inline]
    pub const fn any(self) -> bool {
        self.mask != 0
    }

    /// Check if no bits are set
    #[inline]
    pub const fn none(self) -> bool {
        self.mask == 0
    }
}

// ============================================================================
// x86_64 Implementations
// ============================================================================

#[cfg(target_arch = "x86_64")]
mod x86_impl {
    use super::*;

    /// Compare two vectors of u8 for equality (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// This function requires SSE2 support. Use runtime detection or compile
    /// with appropriate target features.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi8(a, b)
    }

    /// Compare two vectors of u8 for equality (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// This function requires AVX2 support.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi8(a, b)
    }

    /// Compare two vectors of u8 for equality (512-bit AVX-512)
    ///
    /// # Safety
    ///
    /// This function requires AVX-512F and AVX-512BW support.
    #[cfg(target_feature = "avx512f")]
    #[target_feature(enable = "avx512f")]
    #[target_feature(enable = "avx512bw")]
    #[inline]
    pub unsafe fn cmp_eq_u8_avx512(a: __m512i, b: __m512i) -> __mmask64 {
        _mm512_cmpeq_epi8_mask(a, b)
    }

    /// Compare two vectors of u16 for equality (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_u16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi16(a, b)
    }

    /// Compare two vectors of u16 for equality (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi16(a, b)
    }

    /// Compare two vectors of u32 for equality (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_u32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi32(a, b)
    }

    /// Compare two vectors of u32 for equality (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u32_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi32(a, b)
    }

    /// Compare two vectors of u64 for equality (128-bit SSE4.1)
    #[target_feature(enable = "sse4.1")]
    #[inline]
    pub unsafe fn cmp_eq_u64_sse41(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi64(a, b)
    }

    /// Compare two vectors of u64 for equality (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u64_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi64(a, b)
    }

    // Greater than comparisons (signed)

    /// Compare two vectors of i8 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi8(a, b)
    }

    /// Compare two vectors of i8 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi8(a, b)
    }

    /// Compare two vectors of i16 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi16(a, b)
    }

    /// Compare two vectors of i16 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi16(a, b)
    }

    /// Compare two vectors of i32 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi32(a, b)
    }

    /// Compare two vectors of i32 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i32_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi32(a, b)
    }

    /// Compare two vectors of i64 for greater than (128-bit SSE4.2)
    #[target_feature(enable = "sse4.2")]
    #[inline]
    pub unsafe fn cmp_gt_i64_sse42(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi64(a, b)
    }

    /// Compare two vectors of i64 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i64_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi64(a, b)
    }

    // Less than comparisons (implemented via greater than with swapped operands)

    /// Compare two vectors of i8 for less than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi8(a, b)
    }

    /// Compare two vectors of i8 for less than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i8_avx2(a: __m256i, b: __m256i) -> __m256i {
        cmp_gt_i8_avx2(b, a)
    }

    /// Compare two vectors of i16 for less than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi16(a, b)
    }

    /// Compare two vectors of i16 for less than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        cmp_gt_i16_avx2(b, a)
    }

    /// Compare two vectors of i32 for less than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi32(a, b)
    }

    /// Compare two vectors of i32 for less than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i32_avx2(a: __m256i, b: __m256i) -> __m256i {
        cmp_gt_i32_avx2(b, a)
    }

    // Unsigned comparisons (requires XOR trick with sign bit)

    /// Compare two vectors of u8 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi8(0x80u8 as i8);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi8(a_signed, b_signed)
    }

    /// Compare two vectors of u8 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        let sign_bit = _mm256_set1_epi8(0x80u8 as i8);
        let a_signed = _mm256_xor_si256(a, sign_bit);
        let b_signed = _mm256_xor_si256(b, sign_bit);
        _mm256_cmpgt_epi8(a_signed, b_signed)
    }

    /// Compare two vectors of u16 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u16_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi16(0x8000u16 as i16);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi16(a_signed, b_signed)
    }

    /// Compare two vectors of u16 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_u16_avx2(a: __m256i, b: __m256i) -> __m256i {
        let sign_bit = _mm256_set1_epi16(0x8000u16 as i16);
        let a_signed = _mm256_xor_si256(a, sign_bit);
        let b_signed = _mm256_xor_si256(b, sign_bit);
        _mm256_cmpgt_epi16(a_signed, b_signed)
    }

    /// Compare two vectors of u32 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u32_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi32(0x80000000u32 as i32);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi32(a_signed, b_signed)
    }

    /// Compare two vectors of u32 for greater than (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_u32_avx2(a: __m256i, b: __m256i) -> __m256i {
        let sign_bit = _mm256_set1_epi32(0x80000000u32 as i32);
        let a_signed = _mm256_xor_si256(a, sign_bit);
        let b_signed = _mm256_xor_si256(b, sign_bit);
        _mm256_cmpgt_epi32(a_signed, b_signed)
    }

    // Floating-point comparisons

    /// Compare two vectors of f32 for equality (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_eq_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpeq_ps(a, b)
    }

    /// Compare two vectors of f32 for equality (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_eq_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_EQ_OQ>(a, b)
    }

    /// Compare two vectors of f32 for not equal (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_ne_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpneq_ps(a, b)
    }

    /// Compare two vectors of f32 for not equal (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ne_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_NEQ_OQ>(a, b)
    }

    /// Compare two vectors of f32 for greater than (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_gt_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpgt_ps(a, b)
    }

    /// Compare two vectors of f32 for greater than (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_gt_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_GT_OQ>(a, b)
    }

    /// Compare two vectors of f32 for greater than or equal (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_ge_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpge_ps(a, b)
    }

    /// Compare two vectors of f32 for greater than or equal (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ge_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_GE_OQ>(a, b)
    }

    /// Compare two vectors of f32 for less than (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_lt_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmplt_ps(a, b)
    }

    /// Compare two vectors of f32 for less than (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_lt_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_LT_OQ>(a, b)
    }

    /// Compare two vectors of f32 for less than or equal (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_le_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmple_ps(a, b)
    }

    /// Compare two vectors of f32 for less than or equal (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_le_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_LE_OQ>(a, b)
    }

    /// Compare two vectors of f64 for equality (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpeq_pd(a, b)
    }

    /// Compare two vectors of f64 for equality (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_eq_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_EQ_OQ>(a, b)
    }

    /// Compare two vectors of f64 for not equal (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_ne_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpneq_pd(a, b)
    }

    /// Compare two vectors of f64 for not equal (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ne_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_NEQ_OQ>(a, b)
    }

    /// Compare two vectors of f64 for greater than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpgt_pd(a, b)
    }

    /// Compare two vectors of f64 for greater than (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_gt_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_GT_OQ>(a, b)
    }

    /// Compare two vectors of f64 for less than (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmplt_pd(a, b)
    }

    /// Compare two vectors of f64 for less than (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_lt_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_LT_OQ>(a, b)
    }

    // Minimum/Maximum operations

    /// Get minimum of two vectors of u8 (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn min_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_min_epu8(a, b)
    }

    /// Get minimum of two vectors of u8 (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn min_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_min_epu8(a, b)
    }

    /// Get maximum of two vectors of u8 (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn max_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_max_epu8(a, b)
    }

    /// Get maximum of two vectors of u8 (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn max_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_max_epu8(a, b)
    }

    /// Get minimum of two vectors of i16 (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn min_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_min_epi16(a, b)
    }

    /// Get minimum of two vectors of i16 (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn min_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_min_epi16(a, b)
    }

    /// Get maximum of two vectors of i16 (128-bit SSE2)
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn max_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_max_epi16(a, b)
    }

    /// Get maximum of two vectors of i16 (256-bit AVX2)
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn max_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_max_epi16(a, b)
    }

    /// Get minimum of two vectors of f32 (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn min_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_min_ps(a, b)
    }

    /// Get minimum of two vectors of f32 (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn min_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_min_ps(a, b)
    }

    /// Get maximum of two vectors of f32 (128-bit SSE)
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn max_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_max_ps(a, b)
    }

    /// Get maximum of two vectors of f32 (256-bit AVX)
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn max_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_max_ps(a, b)
    }
}

// ============================================================================
// aarch64 Implementations
// ============================================================================

#[cfg(target_arch = "aarch64")]
mod aarch64_impl {
    //! NEON SIMD implementations for aarch64
    //!
    //! # Safety
    //!
    //! All functions in this module require NEON support, which is available on all
    //! aarch64 platforms. The caller must ensure that input pointers are valid and
    //! properly aligned for SIMD operations.

    #![allow(clippy::missing_safety_doc)]

    use super::*;

    // NEON implementations (128-bit)

    /// Compare two vectors of u8 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_u8(a, b) }
    }

    /// Compare two vectors of u16 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_u16(a, b) }
    }

    /// Compare two vectors of u32 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_u32(a, b) }
    }

    /// Compare two vectors of u64 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_u64_neon(a: uint64x2_t, b: uint64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_u64(a, b) }
    }

    /// Compare two vectors of i8 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_i8_neon(a: int8x16_t, b: int8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_s8(a, b) }
    }

    /// Compare two vectors of i16 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_i16_neon(a: int16x8_t, b: int16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_s16(a, b) }
    }

    /// Compare two vectors of i32 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_i32_neon(a: int32x4_t, b: int32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_s32(a, b) }
    }

    /// Compare two vectors of i64 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_i64_neon(a: int64x2_t, b: int64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_s64(a, b) }
    }

    /// Compare two vectors of f32 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_f32_neon(a: float32x4_t, b: float32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_f32(a, b) }
    }

    /// Compare two vectors of f64 for equality (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_eq_f64_neon(a: float64x2_t, b: float64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vceqq_f64(a, b) }
    }

    // Greater than comparisons

    /// Compare two vectors of u8 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_u8(a, b) }
    }

    /// Compare two vectors of u16 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_u16(a, b) }
    }

    /// Compare two vectors of u32 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_u32(a, b) }
    }

    /// Compare two vectors of u64 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_u64_neon(a: uint64x2_t, b: uint64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_u64(a, b) }
    }

    /// Compare two vectors of i8 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_i8_neon(a: int8x16_t, b: int8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_s8(a, b) }
    }

    /// Compare two vectors of i16 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_i16_neon(a: int16x8_t, b: int16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_s16(a, b) }
    }

    /// Compare two vectors of i32 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_i32_neon(a: int32x4_t, b: int32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_s32(a, b) }
    }

    /// Compare two vectors of i64 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_i64_neon(a: int64x2_t, b: int64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_s64(a, b) }
    }

    /// Compare two vectors of f32 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_f32_neon(a: float32x4_t, b: float32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_f32(a, b) }
    }

    /// Compare two vectors of f64 for greater than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_gt_f64_neon(a: float64x2_t, b: float64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgtq_f64(a, b) }
    }

    // Greater than or equal comparisons

    /// Compare two vectors of u8 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_u8(a, b) }
    }

    /// Compare two vectors of u16 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_u16(a, b) }
    }

    /// Compare two vectors of u32 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_u32(a, b) }
    }

    /// Compare two vectors of u64 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_u64_neon(a: uint64x2_t, b: uint64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_u64(a, b) }
    }

    /// Compare two vectors of i8 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_i8_neon(a: int8x16_t, b: int8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_s8(a, b) }
    }

    /// Compare two vectors of i16 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_i16_neon(a: int16x8_t, b: int16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_s16(a, b) }
    }

    /// Compare two vectors of i32 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_i32_neon(a: int32x4_t, b: int32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_s32(a, b) }
    }

    /// Compare two vectors of i64 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_i64_neon(a: int64x2_t, b: int64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_s64(a, b) }
    }

    /// Compare two vectors of f32 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_f32_neon(a: float32x4_t, b: float32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_f32(a, b) }
    }

    /// Compare two vectors of f64 for greater than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_ge_f64_neon(a: float64x2_t, b: float64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcgeq_f64(a, b) }
    }

    // Less than comparisons

    /// Compare two vectors of u8 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_u8(a, b) }
    }

    /// Compare two vectors of u16 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_u16(a, b) }
    }

    /// Compare two vectors of u32 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_u32(a, b) }
    }

    /// Compare two vectors of u64 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_u64_neon(a: uint64x2_t, b: uint64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_u64(a, b) }
    }

    /// Compare two vectors of i8 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_i8_neon(a: int8x16_t, b: int8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_s8(a, b) }
    }

    /// Compare two vectors of i16 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_i16_neon(a: int16x8_t, b: int16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_s16(a, b) }
    }

    /// Compare two vectors of i32 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_i32_neon(a: int32x4_t, b: int32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_s32(a, b) }
    }

    /// Compare two vectors of i64 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_i64_neon(a: int64x2_t, b: int64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_s64(a, b) }
    }

    /// Compare two vectors of f32 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_f32_neon(a: float32x4_t, b: float32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_f32(a, b) }
    }

    /// Compare two vectors of f64 for less than (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_lt_f64_neon(a: float64x2_t, b: float64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcltq_f64(a, b) }
    }

    // Less than or equal comparisons

    /// Compare two vectors of u8 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_u8(a, b) }
    }

    /// Compare two vectors of u16 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_u16(a, b) }
    }

    /// Compare two vectors of u32 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_u32(a, b) }
    }

    /// Compare two vectors of u64 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_u64_neon(a: uint64x2_t, b: uint64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_u64(a, b) }
    }

    /// Compare two vectors of i8 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_i8_neon(a: int8x16_t, b: int8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_s8(a, b) }
    }

    /// Compare two vectors of i16 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_i16_neon(a: int16x8_t, b: int16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_s16(a, b) }
    }

    /// Compare two vectors of i32 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_i32_neon(a: int32x4_t, b: int32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_s32(a, b) }
    }

    /// Compare two vectors of i64 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_i64_neon(a: int64x2_t, b: int64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_s64(a, b) }
    }

    /// Compare two vectors of f32 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_f32_neon(a: float32x4_t, b: float32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_f32(a, b) }
    }

    /// Compare two vectors of f64 for less than or equal (128-bit NEON)
    #[inline]
    pub unsafe fn cmp_le_f64_neon(a: float64x2_t, b: float64x2_t) -> uint64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vcleq_f64(a, b) }
    }

    // Minimum/Maximum operations

    /// Get minimum of two vectors of u8 (128-bit NEON)
    #[inline]
    pub unsafe fn min_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_u8(a, b) }
    }

    /// Get maximum of two vectors of u8 (128-bit NEON)
    #[inline]
    pub unsafe fn max_u8_neon(a: uint8x16_t, b: uint8x16_t) -> uint8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_u8(a, b) }
    }

    /// Get minimum of two vectors of u16 (128-bit NEON)
    #[inline]
    pub unsafe fn min_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_u16(a, b) }
    }

    /// Get maximum of two vectors of u16 (128-bit NEON)
    #[inline]
    pub unsafe fn max_u16_neon(a: uint16x8_t, b: uint16x8_t) -> uint16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_u16(a, b) }
    }

    /// Get minimum of two vectors of u32 (128-bit NEON)
    #[inline]
    pub unsafe fn min_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_u32(a, b) }
    }

    /// Get maximum of two vectors of u32 (128-bit NEON)
    #[inline]
    pub unsafe fn max_u32_neon(a: uint32x4_t, b: uint32x4_t) -> uint32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_u32(a, b) }
    }

    /// Get minimum of two vectors of i8 (128-bit NEON)
    #[inline]
    pub unsafe fn min_i8_neon(a: int8x16_t, b: int8x16_t) -> int8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_s8(a, b) }
    }

    /// Get maximum of two vectors of i8 (128-bit NEON)
    #[inline]
    pub unsafe fn max_i8_neon(a: int8x16_t, b: int8x16_t) -> int8x16_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_s8(a, b) }
    }

    /// Get minimum of two vectors of i16 (128-bit NEON)
    #[inline]
    pub unsafe fn min_i16_neon(a: int16x8_t, b: int16x8_t) -> int16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_s16(a, b) }
    }

    /// Get maximum of two vectors of i16 (128-bit NEON)
    #[inline]
    pub unsafe fn max_i16_neon(a: int16x8_t, b: int16x8_t) -> int16x8_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_s16(a, b) }
    }

    /// Get minimum of two vectors of i32 (128-bit NEON)
    #[inline]
    pub unsafe fn min_i32_neon(a: int32x4_t, b: int32x4_t) -> int32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_s32(a, b) }
    }

    /// Get maximum of two vectors of i32 (128-bit NEON)
    #[inline]
    pub unsafe fn max_i32_neon(a: int32x4_t, b: int32x4_t) -> int32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_s32(a, b) }
    }

    /// Get minimum of two vectors of f32 (128-bit NEON)
    #[inline]
    pub unsafe fn min_f32_neon(a: float32x4_t, b: float32x4_t) -> float32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_f32(a, b) }
    }

    /// Get maximum of two vectors of f32 (128-bit NEON)
    #[inline]
    pub unsafe fn max_f32_neon(a: float32x4_t, b: float32x4_t) -> float32x4_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_f32(a, b) }
    }

    /// Get minimum of two vectors of f64 (128-bit NEON)
    #[inline]
    pub unsafe fn min_f64_neon(a: float64x2_t, b: float64x2_t) -> float64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vminq_f64(a, b) }
    }

    /// Get maximum of two vectors of f64 (128-bit NEON)
    #[inline]
    pub unsafe fn max_f64_neon(a: float64x2_t, b: float64x2_t) -> float64x2_t {
        // SAFETY: Caller ensures NEON support
        unsafe { vmaxq_f64(a, b) }
    }
}

// Re-export platform-specific implementations
#[cfg(target_arch = "x86_64")]
pub use x86_impl::*;

#[cfg(target_arch = "aarch64")]
pub use aarch64_impl::*;

// ============================================================================
// High-level API - Runtime feature detection
// ============================================================================

/// Compare two slices of u8 for equality
///
/// This function automatically selects the best available SIMD implementation
/// based on runtime CPU feature detection.
///
/// # Examples
///
/// ```rust
/// use litchi::common::simd::cmp::simd_eq_u8;
///
/// let a = vec![1u8, 2, 3, 4];
/// let b = vec![1u8, 2, 0, 4];
/// let mut result = vec![0u8; 4];
///
/// simd_eq_u8(&a, &b, &mut result);
/// // result contains: [255, 255, 0, 255]
/// ```
pub fn simd_eq_u8(a: &[u8], b: &[u8], result: &mut [u8]) {
    assert_eq!(a.len(), b.len());
    assert_eq!(a.len(), result.len());

    #[cfg(target_arch = "x86_64")]
    {
        #[cfg(target_feature = "avx2")]
        {
            simd_eq_u8_avx2_impl(a, b, result);
            return;
        }

        if is_x86_feature_detected!("avx2") {
            unsafe { simd_eq_u8_avx2_impl(a, b, result) };
        } else if is_x86_feature_detected!("sse2") {
            unsafe { simd_eq_u8_sse2_impl(a, b, result) };
        } else {
            simd_eq_u8_scalar(a, b, result);
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
        unsafe { simd_eq_u8_neon_impl(a, b, result) };
    }

    #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
    {
        simd_eq_u8_scalar(a, b, result);
    }
}

/// Compare two slices of u8 for inequality
pub fn simd_ne_u8(a: &[u8], b: &[u8], result: &mut [u8]) {
    simd_eq_u8(a, b, result);
    // Invert the result
    for r in result.iter_mut() {
        *r = !*r;
    }
}

/// Check if all bytes in a slice are zero (SIMD-optimized)
///
/// Uses SIMD movemask instructions to avoid scalar iteration.
///
/// # Examples
///
/// ```rust
/// use litchi::common::simd::cmp::is_all_zero;
///
/// let zeros = [0u8; 16];
/// assert!(is_all_zero(&zeros));
///
/// let not_zeros = [0u8, 1, 0, 0, 0, 0, 0, 0];
/// assert!(!is_all_zero(&not_zeros));
/// ```
#[inline]
pub fn is_all_zero(bytes: &[u8]) -> bool {
    #[cfg(target_arch = "x86_64")]
    {
        if is_x86_feature_detected!("avx2") {
            unsafe { is_all_zero_avx2(bytes) }
        } else if is_x86_feature_detected!("sse2") {
            unsafe { is_all_zero_sse2(bytes) }
        } else {
            is_all_zero_scalar(bytes)
        }
    }

    #[cfg(target_arch = "aarch64")]
    {
        unsafe { is_all_zero_neon(bytes) }
    }

    #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
    {
        is_all_zero_scalar(bytes)
    }
}

// Scalar fallback
#[inline]
#[allow(dead_code)]
fn is_all_zero_scalar(bytes: &[u8]) -> bool {
    bytes.iter().all(|&b| b == 0)
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "sse2")]
unsafe fn is_all_zero_sse2(bytes: &[u8]) -> bool {
    let mut i = 0;
    let len = bytes.len();

    // Process 16 bytes at a time
    while i + 16 <= len {
        let vec = unsafe { _mm_loadu_si128(bytes.as_ptr().add(i) as *const __m128i) };
        let zero = unsafe { _mm_setzero_si128() };
        let eq = unsafe { _mm_cmpeq_epi8(vec, zero) };
        let mask = unsafe { _mm_movemask_epi8(eq) };

        // If not all bytes are zero, mask will not be 0xFFFF
        if mask != 0xFFFF {
            return false;
        }
        i += 16;
    }

    // Check remaining bytes
    for &byte in &bytes[i..] {
        if byte != 0 {
            return false;
        }
    }

    true
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "avx2")]
unsafe fn is_all_zero_avx2(bytes: &[u8]) -> bool {
    let mut i = 0;
    let len = bytes.len();

    // Process 32 bytes at a time
    while i + 32 <= len {
        let vec = unsafe { _mm256_loadu_si256(bytes.as_ptr().add(i) as *const __m256i) };
        let zero = unsafe { _mm256_setzero_si256() };
        let eq = unsafe { _mm256_cmpeq_epi8(vec, zero) };
        let mask = unsafe { _mm256_movemask_epi8(eq) };

        // If not all bytes are zero, mask will not be 0xFFFFFFFF
        if mask != -1 {
            // -1 as i32 is 0xFFFFFFFF
            return false;
        }
        i += 32;
    }

    // Handle remaining bytes with SSE2
    if i + 16 <= len {
        let vec = unsafe { _mm_loadu_si128(bytes.as_ptr().add(i) as *const __m128i) };
        let zero = unsafe { _mm_setzero_si128() };
        let eq = unsafe { _mm_cmpeq_epi8(vec, zero) };
        let mask = unsafe { _mm_movemask_epi8(eq) };

        if mask != 0xFFFF {
            return false;
        }
        i += 16;
    }

    // Check remaining bytes
    for &byte in &bytes[i..] {
        if byte != 0 {
            return false;
        }
    }

    true
}

#[cfg(target_arch = "aarch64")]
unsafe fn is_all_zero_neon(bytes: &[u8]) -> bool {
    let mut i = 0;
    let len = bytes.len();

    // Process 16 bytes at a time
    while i + 16 <= len {
        // SAFETY: We've checked bounds
        unsafe {
            let vec = vld1q_u8(bytes.as_ptr().add(i));
            let zero = vdupq_n_u8(0);
            let eq = vceqq_u8(vec, zero);

            // Check if all lanes are 0xFF (all bits set)
            // Use vminvq_u8 to get the minimum value across all lanes
            // If all are 0xFF, min will be 0xFF
            let min = vminvq_u8(eq);
            if min != 0xFF {
                return false;
            }
        }
        i += 16;
    }

    // Check remaining bytes
    for &byte in &bytes[i..] {
        if byte != 0 {
            return false;
        }
    }

    true
}

// Scalar fallback implementations

#[inline]
#[allow(dead_code)] // Used as fallback on non-SIMD platforms
fn simd_eq_u8_scalar(a: &[u8], b: &[u8], result: &mut [u8]) {
    for i in 0..a.len() {
        result[i] = if a[i] == b[i] { 0xFF } else { 0x00 };
    }
}

// Platform-specific implementations for high-level API

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "sse2")]
unsafe fn simd_eq_u8_sse2_impl(a: &[u8], b: &[u8], result: &mut [u8]) {
    let len = a.len();
    let mut i = 0;

    // Process 16 bytes at a time
    while i + 16 <= len {
        let va = _mm_loadu_si128(a.as_ptr().add(i) as *const __m128i);
        let vb = _mm_loadu_si128(b.as_ptr().add(i) as *const __m128i);
        let vcmp = cmp_eq_u8_sse2(va, vb);
        _mm_storeu_si128(result.as_mut_ptr().add(i) as *mut __m128i, vcmp);
        i += 16;
    }

    // Handle remaining elements
    for j in i..len {
        result[j] = if a[j] == b[j] { 0xFF } else { 0x00 };
    }
}

#[cfg(target_arch = "x86_64")]
#[target_feature(enable = "avx2")]
unsafe fn simd_eq_u8_avx2_impl(a: &[u8], b: &[u8], result: &mut [u8]) {
    let len = a.len();
    let mut i = 0;

    // Process 32 bytes at a time
    while i + 32 <= len {
        let va = _mm256_loadu_si256(a.as_ptr().add(i) as *const __m256i);
        let vb = _mm256_loadu_si256(b.as_ptr().add(i) as *const __m256i);
        let vcmp = cmp_eq_u8_avx2(va, vb);
        _mm256_storeu_si256(result.as_mut_ptr().add(i) as *mut __m256i, vcmp);
        i += 32;
    }

    // Handle remaining elements with SSE2
    if i + 16 <= len {
        let va = _mm_loadu_si128(a.as_ptr().add(i) as *const __m128i);
        let vb = _mm_loadu_si128(b.as_ptr().add(i) as *const __m128i);
        let vcmp = cmp_eq_u8_sse2(va, vb);
        _mm_storeu_si128(result.as_mut_ptr().add(i) as *mut __m128i, vcmp);
        i += 16;
    }

    // Handle remaining elements
    for j in i..len {
        result[j] = if a[j] == b[j] { 0xFF } else { 0x00 };
    }
}

#[cfg(target_arch = "aarch64")]
unsafe fn simd_eq_u8_neon_impl(a: &[u8], b: &[u8], result: &mut [u8]) {
    let len = a.len();
    let mut i = 0;

    // Process 16 bytes at a time
    while i + 16 <= len {
        // SAFETY: Caller ensures NEON support and valid memory access
        let va = unsafe { vld1q_u8(a.as_ptr().add(i)) };
        let vb = unsafe { vld1q_u8(b.as_ptr().add(i)) };
        let vcmp = unsafe { cmp_eq_u8_neon(va, vb) };
        unsafe { vst1q_u8(result.as_mut_ptr().add(i), vcmp) };
        i += 16;
    }

    // Handle remaining elements
    for j in i..len {
        result[j] = if a[j] == b[j] { 0xFF } else { 0x00 };
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simd_eq_u8() {
        let a = vec![1u8, 2, 3, 4, 5, 6, 7, 8];
        let b = vec![1u8, 2, 0, 4, 5, 0, 7, 8];
        let mut result = vec![0u8; 8];

        simd_eq_u8(&a, &b, &mut result);

        assert_eq!(result[0], 0xFF);
        assert_eq!(result[1], 0xFF);
        assert_eq!(result[2], 0x00);
        assert_eq!(result[3], 0xFF);
        assert_eq!(result[4], 0xFF);
        assert_eq!(result[5], 0x00);
        assert_eq!(result[6], 0xFF);
        assert_eq!(result[7], 0xFF);
    }

    #[test]
    fn test_simd_ne_u8() {
        let a = vec![1u8, 2, 3, 4];
        let b = vec![1u8, 2, 0, 4];
        let mut result = vec![0u8; 4];

        simd_ne_u8(&a, &b, &mut result);

        assert_eq!(result[0], 0x00);
        assert_eq!(result[1], 0x00);
        assert_eq!(result[2], 0xFF);
        assert_eq!(result[3], 0x00);
    }

    #[test]
    fn test_scalar_fallback() {
        let a = vec![10u8, 20, 30, 40];
        let b = vec![10u8, 25, 30, 45];
        let mut result = vec![0u8; 4];

        simd_eq_u8_scalar(&a, &b, &mut result);

        assert_eq!(result[0], 0xFF);
        assert_eq!(result[1], 0x00);
        assert_eq!(result[2], 0xFF);
        assert_eq!(result[3], 0x00);
    }

    #[test]
    fn test_is_all_zero() {
        // Test all zeros
        let zeros = [0u8; 16];
        assert!(is_all_zero(&zeros));

        // Test with one non-zero byte at different positions
        let mut data = [0u8; 16];
        data[0] = 1;
        assert!(!is_all_zero(&data));

        data[0] = 0;
        data[8] = 1;
        assert!(!is_all_zero(&data));

        data[8] = 0;
        data[15] = 1;
        assert!(!is_all_zero(&data));

        // Test with all non-zero
        let all_ones = [0xFFu8; 16];
        assert!(!is_all_zero(&all_ones));

        // Test different sizes
        assert!(is_all_zero(&[0u8; 8]));
        assert!(is_all_zero(&[0u8; 32]));
        assert!(is_all_zero(&[0u8; 64]));

        assert!(!is_all_zero(&[1u8]));
        assert!(is_all_zero(&[0u8]));

        // Test edge case with size not aligned to SIMD width
        let mut unaligned = vec![0u8; 17];
        assert!(is_all_zero(&unaligned));
        unaligned[16] = 1;
        assert!(!is_all_zero(&unaligned));
    }
}
