//! SIMD vector comparison operations
//!
//! This module provides high-performance SIMD comparison operations for various
//! architectures (x86_64, aarch64) and instruction sets.
//!
//! # Supported Instruction Sets
//!
//! ## x86_64
//! - **SSE2**: 128-bit integer operations
//! - **SSE4.1**: Enhanced 128-bit operations
//! - **SSE4.2**: 128-bit operations with string processing
//! - **AVX2**: 256-bit integer operations
//! - **AVX-512F/BW**: 512-bit operations
//!
//! ## aarch64 (ARM)
//! - **NEON**: Fixed 128-bit SIMD operations (always available)
//! - **SVE**: Scalable Vector Extension with 128-2048 bit vectors
//! - **SVE2**: Enhanced SVE with additional DSP and multimedia instructions
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
//!
//! # SVE/SVE2 Support
//!
//! SVE (Scalable Vector Extension) provides variable-length vectors from 128 to 2048 bits,
//! determined at runtime based on the hardware. This allows the same code to automatically
//! take advantage of larger vector registers when available.
//!
//! Key features of SVE:
//! - **Scalable vectors**: Vector length is runtime-determined
//! - **Predicate-based**: All operations use predicate masks for conditional execution
//! - **Loop-friendly**: Built-in support for processing variable-length data
//!
//! SVE2 extends SVE with additional instructions for:
//! - DSP and multimedia processing
//! - Complex number arithmetic
//! - Additional bit manipulation operations
//! - Histogram and table lookup operations

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
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
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
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
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
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_u16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi16(a, b)
    }

    /// Compare two vectors of u16 for equality (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi16(a, b)
    }

    /// Compare two vectors of u32 for equality (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_u32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi32(a, b)
    }

    /// Compare two vectors of u32 for equality (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u32_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi32(a, b)
    }

    /// Compare two vectors of u64 for equality (128-bit SSE4.1)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE4.1 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE4.1 instructions are available on the target CPU.
    #[target_feature(enable = "sse4.1")]
    #[inline]
    pub unsafe fn cmp_eq_u64_sse41(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpeq_epi64(a, b)
    }

    /// Compare two vectors of u64 for equality (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_eq_u64_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpeq_epi64(a, b)
    }

    // Greater than comparisons (signed)

    /// Compare two vectors of i8 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi8(a, b)
    }

    /// Compare two vectors of i8 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi8(a, b)
    }

    /// Compare two vectors of i16 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi16(a, b)
    }

    /// Compare two vectors of i16 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi16(a, b)
    }

    /// Compare two vectors of i32 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_i32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi32(a, b)
    }

    /// Compare two vectors of i32 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i32_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi32(a, b)
    }

    /// Compare two vectors of i64 for greater than (128-bit SSE4.2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE4.2 instructions are available on the target CPU.
    #[target_feature(enable = "sse4.2")]
    #[inline]
    pub unsafe fn cmp_gt_i64_sse42(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmpgt_epi64(a, b)
    }

    /// Compare two vectors of i64 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_i64_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_cmpgt_epi64(a, b)
    }

    // Less than comparisons (implemented via greater than with swapped operands)

    /// Compare two vectors of i8 for less than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi8(a, b)
    }

    /// Compare two vectors of i8 for less than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i8_avx2(a: __m256i, b: __m256i) -> __m256i {
        unsafe { cmp_gt_i8_avx2(b, a) }
    }

    /// Compare two vectors of i16 for less than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi16(a, b)
    }

    /// Compare two vectors of i16 for less than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        unsafe { cmp_gt_i16_avx2(b, a) }
    }

    /// Compare two vectors of i32 for less than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_i32_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_cmplt_epi32(a, b)
    }

    /// Compare two vectors of i32 for less than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_lt_i32_avx2(a: __m256i, b: __m256i) -> __m256i {
        unsafe { cmp_gt_i32_avx2(b, a) }
    }

    // Unsigned comparisons (requires XOR trick with sign bit)

    /// Compare two vectors of u8 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi8(0x80u8 as i8);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi8(a_signed, b_signed)
    }

    /// Compare two vectors of u8 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        let sign_bit = _mm256_set1_epi8(0x80u8 as i8);
        let a_signed = _mm256_xor_si256(a, sign_bit);
        let b_signed = _mm256_xor_si256(b, sign_bit);
        _mm256_cmpgt_epi8(a_signed, b_signed)
    }

    /// Compare two vectors of u16 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u16_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi16(0x8000u16 as i16);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi16(a_signed, b_signed)
    }

    /// Compare two vectors of u16 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn cmp_gt_u16_avx2(a: __m256i, b: __m256i) -> __m256i {
        let sign_bit = _mm256_set1_epi16(0x8000u16 as i16);
        let a_signed = _mm256_xor_si256(a, sign_bit);
        let b_signed = _mm256_xor_si256(b, sign_bit);
        _mm256_cmpgt_epi16(a_signed, b_signed)
    }

    /// Compare two vectors of u32 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_u32_sse2(a: __m128i, b: __m128i) -> __m128i {
        let sign_bit = _mm_set1_epi32(0x80000000u32 as i32);
        let a_signed = _mm_xor_si128(a, sign_bit);
        let b_signed = _mm_xor_si128(b, sign_bit);
        _mm_cmpgt_epi32(a_signed, b_signed)
    }

    /// Compare two vectors of u32 for greater than (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
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
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_eq_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpeq_ps(a, b)
    }

    /// Compare two vectors of f32 for equality (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_eq_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_EQ_OQ>(a, b)
    }

    /// Compare two vectors of f32 for not equal (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_ne_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpneq_ps(a, b)
    }

    /// Compare two vectors of f32 for not equal (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ne_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_NEQ_OQ>(a, b)
    }

    /// Compare two vectors of f32 for greater than (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_gt_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpgt_ps(a, b)
    }

    /// Compare two vectors of f32 for greater than (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_gt_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_GT_OQ>(a, b)
    }

    /// Compare two vectors of f32 for greater than or equal (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_ge_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmpge_ps(a, b)
    }

    /// Compare two vectors of f32 for greater than or equal (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ge_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_GE_OQ>(a, b)
    }

    /// Compare two vectors of f32 for less than (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_lt_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmplt_ps(a, b)
    }

    /// Compare two vectors of f32 for less than (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_lt_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_LT_OQ>(a, b)
    }

    /// Compare two vectors of f32 for less than or equal (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn cmp_le_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_cmple_ps(a, b)
    }

    /// Compare two vectors of f32 for less than or equal (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_le_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_cmp_ps::<_CMP_LE_OQ>(a, b)
    }

    /// Compare two vectors of f64 for equality (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_eq_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpeq_pd(a, b)
    }

    /// Compare two vectors of f64 for equality (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_eq_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_EQ_OQ>(a, b)
    }

    /// Compare two vectors of f64 for not equal (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_ne_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpneq_pd(a, b)
    }

    /// Compare two vectors of f64 for not equal (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_ne_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_NEQ_OQ>(a, b)
    }

    /// Compare two vectors of f64 for greater than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_gt_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmpgt_pd(a, b)
    }

    /// Compare two vectors of f64 for greater than (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_gt_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_GT_OQ>(a, b)
    }

    /// Compare two vectors of f64 for less than (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn cmp_lt_f64_sse2(a: __m128d, b: __m128d) -> __m128d {
        _mm_cmplt_pd(a, b)
    }

    /// Compare two vectors of f64 for less than (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn cmp_lt_f64_avx(a: __m256d, b: __m256d) -> __m256d {
        _mm256_cmp_pd::<_CMP_LT_OQ>(a, b)
    }

    // Minimum/Maximum operations

    /// Get minimum of two vectors of u8 (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn min_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_min_epu8(a, b)
    }

    /// Get minimum of two vectors of u8 (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn min_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_min_epu8(a, b)
    }

    /// Get maximum of two vectors of u8 (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn max_u8_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_max_epu8(a, b)
    }

    /// Get maximum of two vectors of u8 (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn max_u8_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_max_epu8(a, b)
    }

    /// Get minimum of two vectors of i16 (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn min_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_min_epi16(a, b)
    }

    /// Get minimum of two vectors of i16 (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn min_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_min_epi16(a, b)
    }

    /// Get maximum of two vectors of i16 (128-bit SSE2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[target_feature(enable = "sse2")]
    #[inline]
    pub unsafe fn max_i16_sse2(a: __m128i, b: __m128i) -> __m128i {
        _mm_max_epi16(a, b)
    }

    /// Get maximum of two vectors of i16 (256-bit AVX2)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[target_feature(enable = "avx2")]
    #[inline]
    pub unsafe fn max_i16_avx2(a: __m256i, b: __m256i) -> __m256i {
        _mm256_max_epi16(a, b)
    }

    /// Get minimum of two vectors of f32 (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn min_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_min_ps(a, b)
    }

    /// Get minimum of two vectors of f32 (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
    #[target_feature(enable = "avx")]
    #[inline]
    pub unsafe fn min_f32_avx(a: __m256, b: __m256) -> __m256 {
        _mm256_min_ps(a, b)
    }

    /// Get maximum of two vectors of f32 (128-bit SSE)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE instructions are available on the target CPU.
    #[target_feature(enable = "sse")]
    #[inline]
    pub unsafe fn max_f32_sse(a: __m128, b: __m128) -> __m128 {
        _mm_max_ps(a, b)
    }

    /// Get maximum of two vectors of f32 (256-bit AVX)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX instructions are available on the target CPU.
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

    // ============================================================================
    // SVE (Scalable Vector Extension) Implementations
    // ============================================================================
    //
    // SVE provides scalable vector lengths from 128 to 2048 bits, determined at runtime.
    // Unlike fixed-width NEON, SVE operations use predicate masks for conditional execution.

    #[cfg(target_feature = "sve")]
    mod sve_impl {
        use super::*;

        /// Compare two SVE vectors of u8 for equality
        ///
        /// # Safety
        ///
        /// Requires SVE support. The vector length is determined at runtime.
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_u8_sve(a: svuint8_t, b: svuint8_t) -> svbool_t {
            // SAFETY: Caller ensures SVE support
            unsafe {
                let pg = svptrue_b8();
                svcmpeq_u8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u16 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_u16_sve(a: svuint16_t, b: svuint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpeq_u16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u32 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_u32_sve(a: svuint32_t, b: svuint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpeq_u32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u64 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_u64_sve(a: svuint64_t, b: svuint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpeq_u64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i8 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_i8_sve(a: svint8_t, b: svint8_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b8();
                svcmpeq_s8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i16 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_i16_sve(a: svint16_t, b: svint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpeq_s16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i32 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_i32_sve(a: svint32_t, b: svint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpeq_s32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i64 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_i64_sve(a: svint64_t, b: svint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpeq_s64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f32 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpeq_f32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f64 for equality
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_eq_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpeq_f64(pg, a, b)
            }
        }

        // Greater than comparisons

        /// Compare two SVE vectors of u8 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_u8_sve(a: svuint8_t, b: svuint8_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b8();
                svcmpgt_u8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u16 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_u16_sve(a: svuint16_t, b: svuint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpgt_u16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u32 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_u32_sve(a: svuint32_t, b: svuint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpgt_u32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u64 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_u64_sve(a: svuint64_t, b: svuint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpgt_u64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i8 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_i8_sve(a: svint8_t, b: svint8_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b8();
                svcmpgt_s8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i16 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_i16_sve(a: svint16_t, b: svint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpgt_s16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i32 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_i32_sve(a: svint32_t, b: svint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpgt_s32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i64 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_i64_sve(a: svint64_t, b: svint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpgt_s64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f32 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpgt_f32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f64 for greater than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_gt_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpgt_f64(pg, a, b)
            }
        }

        // Greater than or equal comparisons

        /// Compare two SVE vectors of u8 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_u8_sve(a: svuint8_t, b: svuint8_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b8();
                svcmpge_u8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u16 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_u16_sve(a: svuint16_t, b: svuint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpge_u16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u32 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_u32_sve(a: svuint32_t, b: svuint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpge_u32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of u64 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_u64_sve(a: svuint64_t, b: svuint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpge_u64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i8 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_i8_sve(a: svint8_t, b: svint8_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b8();
                svcmpge_s8(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i16 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_i16_sve(a: svint16_t, b: svint16_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b16();
                svcmpge_s16(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i32 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_i32_sve(a: svint32_t, b: svint32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpge_s32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of i64 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_i64_sve(a: svint64_t, b: svint64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpge_s64(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f32 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b32();
                svcmpge_f32(pg, a, b)
            }
        }

        /// Compare two SVE vectors of f64 for greater than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_ge_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svbool_t {
            unsafe {
                let pg = svptrue_b64();
                svcmpge_f64(pg, a, b)
            }
        }

        // Less than comparisons (implemented via greater than with swapped operands)

        /// Compare two SVE vectors of u8 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_u8_sve(a: svuint8_t, b: svuint8_t) -> svbool_t {
            unsafe { cmp_gt_u8_sve(b, a) }
        }

        /// Compare two SVE vectors of u16 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_u16_sve(a: svuint16_t, b: svuint16_t) -> svbool_t {
            unsafe { cmp_gt_u16_sve(b, a) }
        }

        /// Compare two SVE vectors of u32 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_u32_sve(a: svuint32_t, b: svuint32_t) -> svbool_t {
            unsafe { cmp_gt_u32_sve(b, a) }
        }

        /// Compare two SVE vectors of u64 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_u64_sve(a: svuint64_t, b: svuint64_t) -> svbool_t {
            unsafe { cmp_gt_u64_sve(b, a) }
        }

        /// Compare two SVE vectors of i8 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_i8_sve(a: svint8_t, b: svint8_t) -> svbool_t {
            unsafe { cmp_gt_i8_sve(b, a) }
        }

        /// Compare two SVE vectors of i16 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_i16_sve(a: svint16_t, b: svint16_t) -> svbool_t {
            unsafe { cmp_gt_i16_sve(b, a) }
        }

        /// Compare two SVE vectors of i32 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_i32_sve(a: svint32_t, b: svint32_t) -> svbool_t {
            unsafe { cmp_gt_i32_sve(b, a) }
        }

        /// Compare two SVE vectors of i64 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_i64_sve(a: svint64_t, b: svint64_t) -> svbool_t {
            unsafe { cmp_gt_i64_sve(b, a) }
        }

        /// Compare two SVE vectors of f32 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svbool_t {
            unsafe { cmp_gt_f32_sve(b, a) }
        }

        /// Compare two SVE vectors of f64 for less than
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_lt_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svbool_t {
            unsafe { cmp_gt_f64_sve(b, a) }
        }

        // Less than or equal comparisons

        /// Compare two SVE vectors of u8 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_u8_sve(a: svuint8_t, b: svuint8_t) -> svbool_t {
            unsafe { cmp_ge_u8_sve(b, a) }
        }

        /// Compare two SVE vectors of u16 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_u16_sve(a: svuint16_t, b: svuint16_t) -> svbool_t {
            unsafe { cmp_ge_u16_sve(b, a) }
        }

        /// Compare two SVE vectors of u32 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_u32_sve(a: svuint32_t, b: svuint32_t) -> svbool_t {
            unsafe { cmp_ge_u32_sve(b, a) }
        }

        /// Compare two SVE vectors of u64 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_u64_sve(a: svuint64_t, b: svuint64_t) -> svbool_t {
            unsafe { cmp_ge_u64_sve(b, a) }
        }

        /// Compare two SVE vectors of i8 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_i8_sve(a: svint8_t, b: svint8_t) -> svbool_t {
            unsafe { cmp_ge_i8_sve(b, a) }
        }

        /// Compare two SVE vectors of i16 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_i16_sve(a: svint16_t, b: svint16_t) -> svbool_t {
            unsafe { cmp_ge_i16_sve(b, a) }
        }

        /// Compare two SVE vectors of i32 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_i32_sve(a: svint32_t, b: svint32_t) -> svbool_t {
            unsafe { cmp_ge_i32_sve(b, a) }
        }

        /// Compare two SVE vectors of i64 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_i64_sve(a: svint64_t, b: svint64_t) -> svbool_t {
            unsafe { cmp_ge_i64_sve(b, a) }
        }

        /// Compare two SVE vectors of f32 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svbool_t {
            unsafe { cmp_ge_f32_sve(b, a) }
        }

        /// Compare two SVE vectors of f64 for less than or equal
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn cmp_le_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svbool_t {
            unsafe { cmp_ge_f64_sve(b, a) }
        }

        // Minimum/Maximum operations

        /// Get minimum of two SVE vectors of u8
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_u8_sve(a: svuint8_t, b: svuint8_t) -> svuint8_t {
            unsafe {
                let pg = svptrue_b8();
                svmin_u8_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of u8
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_u8_sve(a: svuint8_t, b: svuint8_t) -> svuint8_t {
            unsafe {
                let pg = svptrue_b8();
                svmax_u8_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of u16
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_u16_sve(a: svuint16_t, b: svuint16_t) -> svuint16_t {
            unsafe {
                let pg = svptrue_b16();
                svmin_u16_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of u16
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_u16_sve(a: svuint16_t, b: svuint16_t) -> svuint16_t {
            unsafe {
                let pg = svptrue_b16();
                svmax_u16_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of u32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_u32_sve(a: svuint32_t, b: svuint32_t) -> svuint32_t {
            unsafe {
                let pg = svptrue_b32();
                svmin_u32_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of u32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_u32_sve(a: svuint32_t, b: svuint32_t) -> svuint32_t {
            unsafe {
                let pg = svptrue_b32();
                svmax_u32_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of i8
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_i8_sve(a: svint8_t, b: svint8_t) -> svint8_t {
            unsafe {
                let pg = svptrue_b8();
                svmin_s8_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of i8
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_i8_sve(a: svint8_t, b: svint8_t) -> svint8_t {
            unsafe {
                let pg = svptrue_b8();
                svmax_s8_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of i16
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_i16_sve(a: svint16_t, b: svint16_t) -> svint16_t {
            unsafe {
                let pg = svptrue_b16();
                svmin_s16_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of i16
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_i16_sve(a: svint16_t, b: svint16_t) -> svint16_t {
            unsafe {
                let pg = svptrue_b16();
                svmax_s16_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of i32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_i32_sve(a: svint32_t, b: svint32_t) -> svint32_t {
            unsafe {
                let pg = svptrue_b32();
                svmin_s32_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of i32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_i32_sve(a: svint32_t, b: svint32_t) -> svint32_t {
            unsafe {
                let pg = svptrue_b32();
                svmax_s32_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of f32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svfloat32_t {
            unsafe {
                let pg = svptrue_b32();
                svmin_f32_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of f32
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_f32_sve(a: svfloat32_t, b: svfloat32_t) -> svfloat32_t {
            unsafe {
                let pg = svptrue_b32();
                svmax_f32_z(pg, a, b)
            }
        }

        /// Get minimum of two SVE vectors of f64
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn min_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svfloat64_t {
            unsafe {
                let pg = svptrue_b64();
                svmin_f64_z(pg, a, b)
            }
        }

        /// Get maximum of two SVE vectors of f64
        #[target_feature(enable = "sve")]
        #[inline]
        pub unsafe fn max_f64_sve(a: svfloat64_t, b: svfloat64_t) -> svfloat64_t {
            unsafe {
                let pg = svptrue_b64();
                svmax_f64_z(pg, a, b)
            }
        }
    }

    #[cfg(target_feature = "sve")]
    pub use sve_impl::*;

    // ============================================================================
    // SVE2 (Scalable Vector Extension 2) Implementations
    // ============================================================================
    //
    // SVE2 builds upon SVE with additional instructions for DSP, multimedia,
    // and general-purpose processing.

    #[cfg(target_feature = "sve2")]
    mod sve2_impl {
        use super::*;

        // SVE2 adds more advanced operations like complex arithmetic, saturating operations,
        // and additional integer operations. For basic comparisons, we can reuse SVE
        // implementations, but SVE2 enables more optimizations for specific use cases.

        /// Check if a byte array contains only ASCII characters using SVE2
        ///
        /// This leverages SVE2's enhanced bit manipulation capabilities.
        #[target_feature(enable = "sve2")]
        #[inline]
        pub unsafe fn is_ascii_sve2(bytes: &[u8]) -> bool {
            unsafe {
                let pg = svptrue_b8();
                let ascii_mask = svdup_u8(0x80);
                let mut i = 0;

                while i < bytes.len() {
                    let remaining = bytes.len() - i;
                    let pg_active = svwhilelt_b8_u64(i as u64, bytes.len() as u64);

                    let data = svld1_u8(pg_active, bytes.as_ptr().add(i));
                    let non_ascii = svtst_u8(pg_active, data, ascii_mask);

                    if svptest_any(pg, non_ascii) {
                        return false;
                    }

                    i += svcntb() as usize;
                }

                true
            }
        }

        /// Count matching bytes using SVE2 histogram operations
        ///
        /// This can be more efficient than NEON for large inputs.
        #[target_feature(enable = "sve2")]
        #[inline]
        pub unsafe fn count_byte_sve2(bytes: &[u8], target: u8) -> usize {
            unsafe {
                let pg = svptrue_b8();
                let target_vec = svdup_u8(target);
                let mut count = 0usize;
                let mut i = 0;

                while i < bytes.len() {
                    let pg_active = svwhilelt_b8_u64(i as u64, bytes.len() as u64);
                    let data = svld1_u8(pg_active, bytes.as_ptr().add(i));
                    let matches = svcmpeq_u8(pg_active, data, target_vec);

                    // Count the number of true predicates
                    count += svcntp_b8(pg, matches) as usize;

                    i += svcntb() as usize;
                }

                count
            }
        }
    }

    #[cfg(target_feature = "sve2")]
    pub use sve2_impl::*;
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
        let zero = _mm_setzero_si128();
        let eq = _mm_cmpeq_epi8(vec, zero);
        let mask = _mm_movemask_epi8(eq);

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
        let zero = _mm256_setzero_si256();
        let eq = _mm256_cmpeq_epi8(vec, zero);
        let mask = _mm256_movemask_epi8(eq);

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
        let zero = _mm_setzero_si128();
        let eq = _mm_cmpeq_epi8(vec, zero);
        let mask = _mm_movemask_epi8(eq);

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

/// Check if all bytes are zero using SVE (scalable vectors)
///
/// This version works with variable-length vectors, making it more efficient
/// on hardware with larger vector registers.
#[cfg(all(target_arch = "aarch64", target_feature = "sve"))]
#[target_feature(enable = "sve")]
unsafe fn is_all_zero_sve(bytes: &[u8]) -> bool {
    unsafe {
        let pg = svptrue_b8();
        let zero_vec = svdup_u8(0);
        let mut i = 0;

        while i < bytes.len() {
            // Create predicate for remaining elements
            let pg_active = svwhilelt_b8_u64(i as u64, bytes.len() as u64);

            // Load vector with predication
            let data = svld1_u8(pg_active, bytes.as_ptr().add(i));

            // Compare with zero
            let eq = svcmpeq_u8(pg_active, data, zero_vec);

            // Check if all active lanes are equal to zero
            // If not all are true, then some byte is non-zero
            if !svptest_last(pg, eq) {
                return false;
            }

            // Advance by the number of bytes processed (vector length)
            i += svcntb() as usize;
        }

        true
    }
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
    unsafe {
        while i + 16 <= len {
            let va = _mm_loadu_si128(a.as_ptr().add(i) as *const __m128i);
            let vb = _mm_loadu_si128(b.as_ptr().add(i) as *const __m128i);
            let vcmp = cmp_eq_u8_sse2(va, vb);
            _mm_storeu_si128(result.as_mut_ptr().add(i) as *mut __m128i, vcmp);
            i += 16;
        }
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

    unsafe {
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

    /// Test SIMD operations with various data sizes to trigger different code paths
    #[test]
    fn test_simd_various_sizes() {
        // Test sizes that trigger different SIMD implementations:
        // - Small (< 16): scalar or SSE
        // - Medium (16-31): SSE2/SSSE3
        // - Large (32+): AVX2
        // - Very large (64+): AVX-512
        let test_sizes = vec![1, 7, 8, 15, 16, 17, 31, 32, 33, 63, 64, 65, 127, 128, 256];

        for size in test_sizes {
            // Create test data
            let a: Vec<u8> = (0..size).map(|i| (i % 256) as u8).collect();
            let mut b = a.clone();

            // Make some elements different
            if size > 4 {
                b[size / 4] = b[size / 4].wrapping_add(1);
                b[size / 2] = b[size / 2].wrapping_add(1);
                b[size * 3 / 4] = b[size * 3 / 4].wrapping_add(1);
            }

            let mut result = vec![0u8; size];

            // Test equality
            simd_eq_u8(&a, &b, &mut result);
            for i in 0..size {
                if a[i] == b[i] {
                    assert_eq!(
                        result[i], 0xFF,
                        "Size {}, index {}: equality mismatch",
                        size, i
                    );
                } else {
                    assert_eq!(
                        result[i], 0x00,
                        "Size {}, index {}: equality mismatch",
                        size, i
                    );
                }
            }

            // Test inequality
            simd_ne_u8(&a, &b, &mut result);
            for i in 0..size {
                if a[i] != b[i] {
                    assert_eq!(
                        result[i], 0xFF,
                        "Size {}, index {}: inequality mismatch",
                        size, i
                    );
                } else {
                    assert_eq!(
                        result[i], 0x00,
                        "Size {}, index {}: inequality mismatch",
                        size, i
                    );
                }
            }
        }
    }

    /// Test edge cases: empty arrays, single elements
    #[test]
    fn test_simd_edge_cases() {
        // Empty arrays
        let empty_a: Vec<u8> = vec![];
        let empty_b: Vec<u8> = vec![];
        let mut empty_result: Vec<u8> = vec![];
        simd_eq_u8(&empty_a, &empty_b, &mut empty_result);
        assert_eq!(empty_result.len(), 0);

        // Single element - equal
        let a = vec![42u8];
        let b = vec![42u8];
        let mut result = vec![0u8];
        simd_eq_u8(&a, &b, &mut result);
        assert_eq!(result[0], 0xFF);

        // Single element - not equal
        let a = vec![42u8];
        let b = vec![43u8];
        let mut result = vec![0u8];
        simd_eq_u8(&a, &b, &mut result);
        assert_eq!(result[0], 0x00);
    }

    /// Test with all identical elements
    #[test]
    fn test_simd_identical_elements() {
        let sizes = vec![8, 16, 32, 64, 100];

        for size in sizes {
            // All equal
            let a = vec![0x55u8; size];
            let b = vec![0x55u8; size];
            let mut result = vec![0u8; size];

            simd_eq_u8(&a, &b, &mut result);
            for &val in &result {
                assert_eq!(val, 0xFF, "All equal test failed for size {}", size);
            }

            // All different
            let a = vec![0x55u8; size];
            let b = vec![0xAAu8; size];
            let mut result = vec![0u8; size];

            simd_eq_u8(&a, &b, &mut result);
            for &val in &result {
                assert_eq!(val, 0x00, "All different test failed for size {}", size);
            }
        }
    }

    /// Test boundary values (0x00, 0xFF, etc.)
    #[test]
    fn test_simd_boundary_values() {
        let test_values = vec![
            (0x00u8, 0x00u8, true),
            (0xFFu8, 0xFFu8, true),
            (0x00u8, 0xFFu8, false),
            (0xFFu8, 0x00u8, false),
            (0x80u8, 0x80u8, true),
            (0x7Fu8, 0x80u8, false),
        ];

        for (val_a, val_b, should_equal) in test_values {
            let a = vec![val_a; 64];
            let b = vec![val_b; 64];
            let mut result = vec![0u8; 64];

            simd_eq_u8(&a, &b, &mut result);

            let expected = if should_equal { 0xFF } else { 0x00 };
            for &val in &result {
                assert_eq!(
                    val, expected,
                    "Boundary value test failed: 0x{:02X} vs 0x{:02X}",
                    val_a, val_b
                );
            }
        }
    }

    /// Test alternating patterns
    #[test]
    fn test_simd_alternating_patterns() {
        let size = 128;
        let mut a = vec![0u8; size];
        let mut b = vec![0u8; size];

        // Create alternating pattern: equal, not-equal, equal, not-equal, ...
        for i in 0..size {
            a[i] = (i % 2) as u8;
            b[i] = if i % 4 < 2 {
                (i % 2) as u8
            } else {
                ((i + 1) % 2) as u8
            };
        }

        let mut result = vec![0u8; size];
        simd_eq_u8(&a, &b, &mut result);

        for i in 0..size {
            let expected = if a[i] == b[i] { 0xFF } else { 0x00 };
            assert_eq!(
                result[i], expected,
                "Alternating pattern failed at index {}",
                i
            );
        }
    }

    /// Test that SIMD and scalar implementations produce the same results
    #[test]
    fn test_simd_vs_scalar_consistency() {
        let test_data = vec![
            (vec![1, 2, 3, 4, 5, 6, 7, 8], vec![1, 2, 3, 4, 5, 6, 7, 8]),
            (vec![0; 64], vec![0; 64]),
            (vec![0xFF; 64], vec![0xFF; 64]),
            (
                (0..100).map(|x| x as u8).collect(),
                (0..100).map(|x| x as u8).collect(),
            ),
            (
                (0..100).map(|x| x as u8).collect(),
                (0..100).map(|x| (x + 1) as u8).collect(),
            ),
        ];

        for (a, b) in test_data {
            let mut simd_result = vec![0u8; a.len()];
            let mut scalar_result = vec![0u8; a.len()];

            simd_eq_u8(&a, &b, &mut simd_result);
            simd_eq_u8_scalar(&a, &b, &mut scalar_result);

            assert_eq!(simd_result, scalar_result, "SIMD and scalar results differ");
        }
    }

    /// Test is_all_zero with various patterns
    #[test]
    fn test_is_all_zero_comprehensive() {
        // Test various sizes
        for size in [
            0, 1, 7, 8, 15, 16, 17, 31, 32, 33, 63, 64, 65, 127, 128, 255, 256,
        ] {
            // All zeros
            assert!(
                is_all_zero(&vec![0u8; size]),
                "is_all_zero failed for size {}",
                size
            );

            // Single non-zero at start
            if size > 0 {
                let mut data = vec![0u8; size];
                data[0] = 1;
                assert!(
                    !is_all_zero(&data),
                    "is_all_zero false positive for size {}",
                    size
                );
            }

            // Single non-zero at end
            if size > 0 {
                let mut data = vec![0u8; size];
                data[size - 1] = 1;
                assert!(
                    !is_all_zero(&data),
                    "is_all_zero false positive for size {}",
                    size
                );
            }

            // Single non-zero in middle
            if size > 1 {
                let mut data = vec![0u8; size];
                data[size / 2] = 1;
                assert!(
                    !is_all_zero(&data),
                    "is_all_zero false positive for size {}",
                    size
                );
            }

            // All non-zero
            if size > 0 {
                assert!(
                    !is_all_zero(&vec![1u8; size]),
                    "is_all_zero false positive for all ones"
                );
            }
        }
    }

    /// Stress test with large random-like data
    #[test]
    fn test_simd_large_data() {
        let size = 1024;
        let a: Vec<u8> = (0..size).map(|i| ((i * 137 + 42) % 256) as u8).collect();
        let mut b = a.clone();

        // Modify every 7th element
        for i in (0..size).step_by(7) {
            b[i] = b[i].wrapping_add(1);
        }

        let mut result = vec![0u8; size];
        simd_eq_u8(&a, &b, &mut result);

        for i in 0..size {
            let expected = if a[i] == b[i] { 0xFF } else { 0x00 };
            assert_eq!(result[i], expected, "Large data test failed at index {}", i);
        }
    }
}
