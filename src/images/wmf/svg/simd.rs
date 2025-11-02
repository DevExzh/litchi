//! SIMD-accelerated point transformation for WMF to SVG conversion
//!
//! Provides batch transformation of multiple points using platform-specific
//! SIMD instructions (AVX/AVX2/AVX512 on x86_64, NEON on aarch64).
//!
//! # Performance
//!
//! - Processes 2-8 points per SIMD operation depending on instruction set
//! - Reduces memory allocations by reusing output buffers
//! - Falls back to scalar operations for remainder points
//!
//! # Supported Platforms
//!
//! - **x86_64**: SSE2, AVX, AVX2, AVX512F
//! - **aarch64**: NEON (always available), SVE (planned)
//! - **Fallback**: Scalar implementation for other architectures

#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::*;

#[cfg(target_arch = "aarch64")]
use std::arch::aarch64::*;

/// SIMD-accelerated coordinate transformer
pub struct SimdTransform {
    bbox_left: f64,
    bbox_top: f64,
    scale_x: f64,
    scale_y: f64,
}

impl SimdTransform {
    /// Create new SIMD transformer with precomputed scale factors
    #[inline]
    pub fn new(bbox_left: f64, bbox_top: f64, scale_x: f64, scale_y: f64) -> Self {
        Self {
            bbox_left,
            bbox_top,
            scale_x,
            scale_y,
        }
    }

    /// Transform multiple points in batch using SIMD (x-coordinates and y-coordinates separately)
    ///
    /// # Arguments
    ///
    /// * `xs` - Input x-coordinates (i16 WMF coordinates)
    /// * `ys` - Input y-coordinates (i16 WMF coordinates)
    /// * `out_x` - Output buffer for transformed x-coordinates (must be >= xs.len())
    /// * `out_y` - Output buffer for transformed y-coordinates (must be >= ys.len())
    ///
    /// # Performance
    ///
    /// Automatically selects best SIMD implementation based on CPU features:
    /// - AVX512F: 8 points per iteration
    /// - AVX/AVX2: 4 points per iteration
    /// - SSE2: 2 points per iteration
    /// - NEON: 2 points per iteration (aarch64)
    /// - Scalar: fallback for remainder and unsupported platforms
    #[inline]
    pub fn transform_batch(&self, xs: &[i16], ys: &[i16], out_x: &mut [f64], out_y: &mut [f64]) {
        let len = xs.len().min(ys.len()).min(out_x.len()).min(out_y.len());

        #[cfg(target_arch = "x86_64")]
        {
            // Runtime feature detection for x86_64
            if is_x86_feature_detected!("avx512f") {
                unsafe { self.transform_batch_avx512(xs, ys, out_x, out_y, len) };
                return;
            }
            if is_x86_feature_detected!("avx") {
                unsafe { self.transform_batch_avx(xs, ys, out_x, out_y, len) };
                return;
            }
            if is_x86_feature_detected!("sse2") {
                unsafe { self.transform_batch_sse2(xs, ys, out_x, out_y, len) };
                return;
            }
            // Fallback to scalar for x86_64 without SSE2
            self.transform_batch_scalar(xs, ys, out_x, out_y, len);
        }

        #[cfg(target_arch = "aarch64")]
        {
            // NEON is always available on aarch64
            unsafe { self.transform_batch_neon(xs, ys, out_x, out_y, len) };
        }

        #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
        {
            // Fallback to scalar for other architectures
            self.transform_batch_scalar(xs, ys, out_x, out_y, len);
        }
    }

    /// Scalar fallback implementation
    #[inline]
    fn transform_batch_scalar(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
        len: usize,
    ) {
        for i in 0..len {
            out_x[i] = (xs[i] as f64 - self.bbox_left) * self.scale_x;
            out_y[i] = (ys[i] as f64 - self.bbox_top) * self.scale_y;
        }
    }

    /// AVX512F implementation (8 points per iteration)
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "avx512f")]
    #[inline]
    unsafe fn transform_batch_avx512(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
        len: usize,
    ) {
        // SAFETY: All AVX-512 intrinsic operations are safe within this target_feature context
        unsafe {
            let chunks = len / 8;
            let remainder = len % 8;

            let bbox_left_vec = _mm512_set1_pd(self.bbox_left);
            let bbox_top_vec = _mm512_set1_pd(self.bbox_top);
            let scale_x_vec = _mm512_set1_pd(self.scale_x);
            let scale_y_vec = _mm512_set1_pd(self.scale_y);

            for i in 0..chunks {
                let idx = i * 8;

                // Load 8 x i16 values and convert to f64
                let x_i32 =
                    _mm256_cvtepi16_epi32(_mm_loadu_si128(xs.as_ptr().add(idx) as *const __m128i));
                let x_f64 = _mm512_cvtepi32_pd(x_i32);

                // Load 8 y i16 values and convert to f64
                let y_i32 =
                    _mm256_cvtepi16_epi32(_mm_loadu_si128(ys.as_ptr().add(idx) as *const __m128i));
                let y_f64 = _mm512_cvtepi32_pd(y_i32);

                // Transform: (coord - bbox_offset) * scale
                let tx = _mm512_mul_pd(_mm512_sub_pd(x_f64, bbox_left_vec), scale_x_vec);
                let ty = _mm512_mul_pd(_mm512_sub_pd(y_f64, bbox_top_vec), scale_y_vec);

                // Store results
                _mm512_storeu_pd(out_x.as_mut_ptr().add(idx), tx);
                _mm512_storeu_pd(out_y.as_mut_ptr().add(idx), ty);
            }

            // Handle remainder with scalar
            let start = chunks * 8;
            self.transform_batch_scalar(
                &xs[start..],
                &ys[start..],
                &mut out_x[start..],
                &mut out_y[start..],
                remainder,
            );
        }
    }

    /// AVX/AVX2 implementation (4 points per iteration)
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "avx")]
    #[inline]
    unsafe fn transform_batch_avx(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
        len: usize,
    ) {
        // SAFETY: All AVX intrinsic operations are safe within this target_feature context
        unsafe {
            let chunks = len / 4;
            let remainder = len % 4;

            let bbox_left_vec = _mm256_set1_pd(self.bbox_left);
            let bbox_top_vec = _mm256_set1_pd(self.bbox_top);
            let scale_x_vec = _mm256_set1_pd(self.scale_x);
            let scale_y_vec = _mm256_set1_pd(self.scale_y);

            for i in 0..chunks {
                let idx = i * 4;

                // Load 4 i16 values, convert to i32, then to f64
                // Read 8 bytes (4 x i16) into lower half of xmm register
                let x_i16 = _mm_loadl_epi64(xs.as_ptr().add(idx) as *const __m128i);
                let x_i32 = _mm_cvtepi16_epi32(x_i16);
                let x_f64 = _mm256_cvtepi32_pd(x_i32);

                let y_i16 = _mm_loadl_epi64(ys.as_ptr().add(idx) as *const __m128i);
                let y_i32 = _mm_cvtepi16_epi32(y_i16);
                let y_f64 = _mm256_cvtepi32_pd(y_i32);

                // Transform: (coord - bbox_offset) * scale
                let tx = _mm256_mul_pd(_mm256_sub_pd(x_f64, bbox_left_vec), scale_x_vec);
                let ty = _mm256_mul_pd(_mm256_sub_pd(y_f64, bbox_top_vec), scale_y_vec);

                // Store results
                _mm256_storeu_pd(out_x.as_mut_ptr().add(idx), tx);
                _mm256_storeu_pd(out_y.as_mut_ptr().add(idx), ty);
            }

            // Handle remainder with scalar
            let start = chunks * 4;
            self.transform_batch_scalar(
                &xs[start..],
                &ys[start..],
                &mut out_x[start..],
                &mut out_y[start..],
                remainder,
            );
        }
    }

    /// SSE2 implementation (2 points per iteration)
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "sse2")]
    #[inline]
    unsafe fn transform_batch_sse2(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
        len: usize,
    ) {
        // SAFETY: All SSE2 intrinsic operations are safe within this target_feature context
        unsafe {
            let chunks = len / 2;
            let remainder = len % 2;

            let bbox_left_vec = _mm_set1_pd(self.bbox_left);
            let bbox_top_vec = _mm_set1_pd(self.bbox_top);
            let scale_x_vec = _mm_set1_pd(self.scale_x);
            let scale_y_vec = _mm_set1_pd(self.scale_y);

            for i in 0..chunks {
                let idx = i * 2;

                // Load 2 i16 values and convert to f64
                let x_f64_1 = _mm_cvtsi32_sd(_mm_setzero_pd(), xs[idx] as i32);
                let x_f64_2 = _mm_cvtsi32_sd(_mm_setzero_pd(), xs[idx + 1] as i32);
                let x_f64 = _mm_unpacklo_pd(x_f64_1, x_f64_2);

                let y_f64_1 = _mm_cvtsi32_sd(_mm_setzero_pd(), ys[idx] as i32);
                let y_f64_2 = _mm_cvtsi32_sd(_mm_setzero_pd(), ys[idx + 1] as i32);
                let y_f64 = _mm_unpacklo_pd(y_f64_1, y_f64_2);

                // Transform: (coord - bbox_offset) * scale
                let tx = _mm_mul_pd(_mm_sub_pd(x_f64, bbox_left_vec), scale_x_vec);
                let ty = _mm_mul_pd(_mm_sub_pd(y_f64, bbox_top_vec), scale_y_vec);

                // Store results
                _mm_storeu_pd(out_x.as_mut_ptr().add(idx), tx);
                _mm_storeu_pd(out_y.as_mut_ptr().add(idx), ty);
            }

            // Handle remainder with scalar
            let start = chunks * 2;
            self.transform_batch_scalar(
                &xs[start..],
                &ys[start..],
                &mut out_x[start..],
                &mut out_y[start..],
                remainder,
            );
        }
    }

    /// NEON implementation for aarch64 (2 points per iteration)
    #[cfg(target_arch = "aarch64")]
    #[target_feature(enable = "neon")]
    #[inline]
    unsafe fn transform_batch_neon(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
        len: usize,
    ) {
        let chunks = len / 2;
        let remainder = len % 2;

        // SAFETY: All NEON operations are safe within this target_feature(enable = "neon") function
        unsafe {
            let bbox_left_vec = vdupq_n_f64(self.bbox_left);
            let bbox_top_vec = vdupq_n_f64(self.bbox_top);
            let scale_x_vec = vdupq_n_f64(self.scale_x);
            let scale_y_vec = vdupq_n_f64(self.scale_y);

            for i in 0..chunks {
                let idx = i * 2;

                // Load 2 i16 values and convert to f64
                let x1 = xs[idx] as f64;
                let x2 = xs[idx + 1] as f64;
                let x_f64 = vld1q_f64([x1, x2].as_ptr());

                let y1 = ys[idx] as f64;
                let y2 = ys[idx + 1] as f64;
                let y_f64 = vld1q_f64([y1, y2].as_ptr());

                // Transform: (coord - bbox_offset) * scale
                let tx = vmulq_f64(vsubq_f64(x_f64, bbox_left_vec), scale_x_vec);
                let ty = vmulq_f64(vsubq_f64(y_f64, bbox_top_vec), scale_y_vec);

                // Store results
                vst1q_f64(out_x.as_mut_ptr().add(idx), tx);
                vst1q_f64(out_y.as_mut_ptr().add(idx), ty);
            }
        }

        // Handle remainder with scalar
        let start = chunks * 2;
        self.transform_batch_scalar(
            &xs[start..],
            &ys[start..],
            &mut out_x[start..],
            &mut out_y[start..],
            remainder,
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simd_transform_batch() {
        let transform = SimdTransform::new(100.0, 200.0, 2.0, 3.0);

        let xs = vec![100i16, 150, 200, 250, 300, 350, 400, 450];
        let ys = vec![200i16, 250, 300, 350, 400, 450, 500, 550];

        let mut out_x = vec![0.0; 8];
        let mut out_y = vec![0.0; 8];

        transform.transform_batch(&xs, &ys, &mut out_x, &mut out_y);

        // Verify results match scalar calculation
        for i in 0..8 {
            let expected_x = (xs[i] as f64 - 100.0) * 2.0;
            let expected_y = (ys[i] as f64 - 200.0) * 3.0;

            assert!((out_x[i] - expected_x).abs() < 1e-10);
            assert!((out_y[i] - expected_y).abs() < 1e-10);
        }
    }

    #[test]
    fn test_simd_transform_non_aligned() {
        let transform = SimdTransform::new(0.0, 0.0, 1.0, 1.0);

        // Test with non-power-of-2 length
        let xs = vec![10i16, 20, 30, 40, 50];
        let ys = vec![100i16, 200, 300, 400, 500];

        let mut out_x = vec![0.0; 5];
        let mut out_y = vec![0.0; 5];

        transform.transform_batch(&xs, &ys, &mut out_x, &mut out_y);

        for i in 0..5 {
            assert_eq!(out_x[i], xs[i] as f64);
            assert_eq!(out_y[i], ys[i] as f64);
        }
    }
}
