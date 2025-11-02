//! Bounding box calculation for WMF files without placeable headers
//!
//! When a WMF file lacks a placeable header, we must scan all drawing records
//! to determine the effective drawing area. This matches libwmf's wmf_scan behavior.
//!
//! # SIMD Acceleration
//!
//! The bounds calculation is accelerated using SIMD instructions when processing
//! records with multiple points (POLYGON, POLYLINE, etc.):
//!
//! - **x86_64**: SSE2, SSE4.1, AVX2, AVX-512
//! - **aarch64**: NEON
//! - **Other**: Scalar fallback

use super::super::constants::record;
use super::super::parser::WmfRecord;
use crate::common::binary::{read_i16_le, read_u16_le};

#[cfg(target_arch = "x86_64")]
use std::arch::x86_64::*;

#[cfg(target_arch = "aarch64")]
use std::arch::aarch64::*;

/// Calculate bounding box by scanning WMF records
pub struct BoundsCalculator;

impl BoundsCalculator {
    /// Scan WMF records to determine bounding box (matches libwmf wmf_scan)
    pub fn scan_records(records: &[WmfRecord]) -> (i16, i16, i16, i16) {
        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;
        let mut font_height = 0i16;

        for rec in records {
            match rec.function {
                record::RECTANGLE | record::ELLIPSE | record::ROUND_RECT
                    if rec.params.len() >= 8 =>
                {
                    let b = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let r = read_i16_le(&rec.params, 2).unwrap_or(0);
                    let t = read_i16_le(&rec.params, 4).unwrap_or(0);
                    let l = read_i16_le(&rec.params, 6).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, l, t);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, r, b);
                },
                record::POLYGON | record::POLYLINE if rec.params.len() >= 2 => {
                    let count = read_i16_le(&rec.params, 0).unwrap_or(0) as usize;
                    let actual_count = count.min((rec.params.len() - 2) / 4);

                    // Use SIMD acceleration for batch processing
                    Self::update_from_points(
                        &mut left,
                        &mut top,
                        &mut right,
                        &mut bottom,
                        &rec.params[2..],
                        actual_count,
                    );
                },
                record::LINE_TO | record::MOVE_TO | record::SET_PIXEL_V
                    if rec.params.len() >= 4 =>
                {
                    let y = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let x = read_i16_le(&rec.params, 2).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                },
                record::ARC | record::PIE | record::CHORD if rec.params.len() >= 16 => {
                    let yend = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let xend = read_i16_le(&rec.params, 2).unwrap_or(0);
                    let ystart = read_i16_le(&rec.params, 4).unwrap_or(0);
                    let xstart = read_i16_le(&rec.params, 6).unwrap_or(0);
                    let b = read_i16_le(&rec.params, 8).unwrap_or(0);
                    let r = read_i16_le(&rec.params, 10).unwrap_or(0);
                    let t = read_i16_le(&rec.params, 12).unwrap_or(0);
                    let l = read_i16_le(&rec.params, 14).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, l, t);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, r, b);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, xstart, ystart);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, xend, yend);
                },
                record::CREATE_FONT_INDIRECT if rec.params.len() >= 2 => {
                    font_height = read_i16_le(&rec.params, 0).unwrap_or(0).abs();
                },
                record::TEXT_OUT if rec.params.len() >= 6 => {
                    let len = read_u16_le(&rec.params, 0).unwrap_or(0) as usize;
                    let off = 2 + len.div_ceil(2) * 2;
                    if rec.params.len() >= off + 4 {
                        let y = read_i16_le(&rec.params, off).unwrap_or(0);
                        let x = read_i16_le(&rec.params, off + 2).unwrap_or(0);
                        Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                        Self::update(
                            &mut left,
                            &mut top,
                            &mut right,
                            &mut bottom,
                            x,
                            y + font_height,
                        );
                    }
                },
                record::EXT_TEXT_OUT if rec.params.len() >= 8 => {
                    let y = read_i16_le(&rec.params, 0).unwrap_or(0);
                    let x = read_i16_le(&rec.params, 2).unwrap_or(0);
                    Self::update(&mut left, &mut top, &mut right, &mut bottom, x, y);
                    Self::update(
                        &mut left,
                        &mut top,
                        &mut right,
                        &mut bottom,
                        x,
                        y + font_height,
                    );
                },
                _ => {},
            }
        }

        // Return bounds or default if empty
        if left != i16::MAX && right != i16::MIN && top != i16::MAX && bottom != i16::MIN {
            (left, top, right, bottom)
        } else {
            (0, 0, 1000, 1000)
        }
    }

    #[inline]
    fn update(left: &mut i16, top: &mut i16, right: &mut i16, bottom: &mut i16, x: i16, y: i16) {
        *left = (*left).min(x);
        *top = (*top).min(y);
        *right = (*right).max(x);
        *bottom = (*bottom).max(y);
    }

    /// Update bounds from a sequence of points with SIMD acceleration
    ///
    /// Points are stored as interleaved x,y coordinates in the params buffer.
    /// Format: [x0_lo, x0_hi, y0_lo, y0_hi, x1_lo, x1_hi, y1_lo, y1_hi, ...]
    fn update_from_points(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        if count == 0 {
            return;
        }

        #[cfg(target_arch = "x86_64")]
        {
            // Runtime feature detection for x86_64
            if is_x86_feature_detected!("avx2") {
                unsafe {
                    Self::update_from_points_avx2(left, top, right, bottom, params, count);
                }
            } else if is_x86_feature_detected!("sse4.1") {
                unsafe {
                    Self::update_from_points_sse41(left, top, right, bottom, params, count);
                }
            } else if is_x86_feature_detected!("sse2") {
                unsafe {
                    Self::update_from_points_sse2(left, top, right, bottom, params, count);
                }
            } else {
                Self::update_from_points_scalar(left, top, right, bottom, params, count);
            }
        }

        #[cfg(target_arch = "aarch64")]
        {
            // NEON is always available on aarch64
            unsafe {
                Self::update_from_points_neon(left, top, right, bottom, params, count);
            }
        }

        // Scalar fallback for other architectures
        #[cfg(not(any(target_arch = "x86_64", target_arch = "aarch64")))]
        {
            Self::update_from_points_scalar(left, top, right, bottom, params, count);
        }
    }

    /// Scalar implementation for processing points
    #[inline]
    fn update_from_points_scalar(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        for i in 0..count {
            let x = read_i16_le(params, i * 4).unwrap_or(0);
            let y = read_i16_le(params, i * 4 + 2).unwrap_or(0);
            Self::update(left, top, right, bottom, x, y);
        }
    }

    /// SSE2 implementation (processes 4 points at a time using scalar loads)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE2 instructions are available on the target CPU.
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "sse2")]
    unsafe fn update_from_points_sse2(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        // SAFETY: All SSE2 intrinsic operations are safe within this target_feature context
        unsafe {
            let mut min_x = _mm_set1_epi16(*left);
            let mut min_y = _mm_set1_epi16(*top);
            let mut max_x = _mm_set1_epi16(*right);
            let mut max_y = _mm_set1_epi16(*bottom);

            let mut i = 0;
            // Process 4 points at a time
            while i + 3 < count && (i * 4 + 16) <= params.len() {
                // Load 4 points as individual i16 values
                let x0 = read_i16_le(params, i * 4).unwrap_or(0);
                let y0 = read_i16_le(params, i * 4 + 2).unwrap_or(0);
                let x1 = read_i16_le(params, (i + 1) * 4).unwrap_or(0);
                let y1 = read_i16_le(params, (i + 1) * 4 + 2).unwrap_or(0);
                let x2 = read_i16_le(params, (i + 2) * 4).unwrap_or(0);
                let y2 = read_i16_le(params, (i + 2) * 4 + 2).unwrap_or(0);
                let x3 = read_i16_le(params, (i + 3) * 4).unwrap_or(0);
                let y3 = read_i16_le(params, (i + 3) * 4 + 2).unwrap_or(0);

                // Pack into SIMD vectors
                let x_vals = _mm_set_epi16(0, 0, 0, 0, x3, x2, x1, x0);
                let y_vals = _mm_set_epi16(0, 0, 0, 0, y3, y2, y1, y0);

                min_x = _mm_min_epi16(min_x, x_vals);
                max_x = _mm_max_epi16(max_x, x_vals);
                min_y = _mm_min_epi16(min_y, y_vals);
                max_y = _mm_max_epi16(max_y, y_vals);

                i += 4;
            }

            // Reduce SIMD results
            let mut temp_min_x = [0i16; 8];
            let mut temp_max_x = [0i16; 8];
            let mut temp_min_y = [0i16; 8];
            let mut temp_max_y = [0i16; 8];
            _mm_storeu_si128(temp_min_x.as_mut_ptr() as *mut __m128i, min_x);
            _mm_storeu_si128(temp_max_x.as_mut_ptr() as *mut __m128i, max_x);
            _mm_storeu_si128(temp_min_y.as_mut_ptr() as *mut __m128i, min_y);
            _mm_storeu_si128(temp_max_y.as_mut_ptr() as *mut __m128i, max_y);

            *left = temp_min_x.iter().copied().min().unwrap_or(*left);
            *right = temp_max_x.iter().copied().max().unwrap_or(*right);
            *top = temp_min_y.iter().copied().min().unwrap_or(*top);
            *bottom = temp_max_y.iter().copied().max().unwrap_or(*bottom);

            // Handle remaining points with scalar code
            Self::update_from_points_scalar(left, top, right, bottom, &params[i * 4..], count - i);
        }
    }

    /// SSE4.1 implementation (processes 8 points at a time using scalar loads)
    ///
    /// # Safety
    ///
    /// Caller must ensure that SSE4.1 instructions are available on the target CPU.
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "sse4.1")]
    unsafe fn update_from_points_sse41(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        // SAFETY: All SSE4.1 intrinsic operations are safe within this target_feature context
        unsafe {
            let mut min_x = _mm_set1_epi16(*left);
            let mut min_y = _mm_set1_epi16(*top);
            let mut max_x = _mm_set1_epi16(*right);
            let mut max_y = _mm_set1_epi16(*bottom);

            let mut i = 0;
            // Process 8 points at a time
            while i + 7 < count && (i * 4 + 32) <= params.len() {
                // Load 8 points as individual i16 values
                let x0 = read_i16_le(params, i * 4).unwrap_or(0);
                let y0 = read_i16_le(params, i * 4 + 2).unwrap_or(0);
                let x1 = read_i16_le(params, (i + 1) * 4).unwrap_or(0);
                let y1 = read_i16_le(params, (i + 1) * 4 + 2).unwrap_or(0);
                let x2 = read_i16_le(params, (i + 2) * 4).unwrap_or(0);
                let y2 = read_i16_le(params, (i + 2) * 4 + 2).unwrap_or(0);
                let x3 = read_i16_le(params, (i + 3) * 4).unwrap_or(0);
                let y3 = read_i16_le(params, (i + 3) * 4 + 2).unwrap_or(0);
                let x4 = read_i16_le(params, (i + 4) * 4).unwrap_or(0);
                let y4 = read_i16_le(params, (i + 4) * 4 + 2).unwrap_or(0);
                let x5 = read_i16_le(params, (i + 5) * 4).unwrap_or(0);
                let y5 = read_i16_le(params, (i + 5) * 4 + 2).unwrap_or(0);
                let x6 = read_i16_le(params, (i + 6) * 4).unwrap_or(0);
                let y6 = read_i16_le(params, (i + 6) * 4 + 2).unwrap_or(0);
                let x7 = read_i16_le(params, (i + 7) * 4).unwrap_or(0);
                let y7 = read_i16_le(params, (i + 7) * 4 + 2).unwrap_or(0);

                // Pack into SIMD vectors
                let x_vals = _mm_set_epi16(x7, x6, x5, x4, x3, x2, x1, x0);
                let y_vals = _mm_set_epi16(y7, y6, y5, y4, y3, y2, y1, y0);

                min_x = _mm_min_epi16(min_x, x_vals);
                max_x = _mm_max_epi16(max_x, x_vals);
                min_y = _mm_min_epi16(min_y, y_vals);
                max_y = _mm_max_epi16(max_y, y_vals);

                i += 8;
            }

            // Reduce SIMD results
            let mut temp_min_x = [0i16; 8];
            let mut temp_max_x = [0i16; 8];
            let mut temp_min_y = [0i16; 8];
            let mut temp_max_y = [0i16; 8];
            _mm_storeu_si128(temp_min_x.as_mut_ptr() as *mut __m128i, min_x);
            _mm_storeu_si128(temp_max_x.as_mut_ptr() as *mut __m128i, max_x);
            _mm_storeu_si128(temp_min_y.as_mut_ptr() as *mut __m128i, min_y);
            _mm_storeu_si128(temp_max_y.as_mut_ptr() as *mut __m128i, max_y);

            *left = temp_min_x.iter().copied().min().unwrap_or(*left);
            *right = temp_max_x.iter().copied().max().unwrap_or(*right);
            *top = temp_min_y.iter().copied().min().unwrap_or(*top);
            *bottom = temp_max_y.iter().copied().max().unwrap_or(*bottom);

            // Handle remaining points with scalar code
            Self::update_from_points_scalar(left, top, right, bottom, &params[i * 4..], count - i);
        }
    }

    /// AVX2 implementation (processes 16 points at a time using scalar loads)
    ///
    /// # Safety
    ///
    /// Caller must ensure that AVX2 instructions are available on the target CPU.
    #[cfg(target_arch = "x86_64")]
    #[target_feature(enable = "avx2")]
    unsafe fn update_from_points_avx2(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        // SAFETY: All AVX2 intrinsic operations are safe within this target_feature context
        unsafe {
            let mut min_x = _mm256_set1_epi16(*left);
            let mut min_y = _mm256_set1_epi16(*top);
            let mut max_x = _mm256_set1_epi16(*right);
            let mut max_y = _mm256_set1_epi16(*bottom);

            let mut i = 0;
            // Process 16 points at a time
            while i + 15 < count && (i * 4 + 64) <= params.len() {
                // Load 16 points as individual i16 values
                let x0 = read_i16_le(params, i * 4).unwrap_or(0);
                let y0 = read_i16_le(params, i * 4 + 2).unwrap_or(0);
                let x1 = read_i16_le(params, (i + 1) * 4).unwrap_or(0);
                let y1 = read_i16_le(params, (i + 1) * 4 + 2).unwrap_or(0);
                let x2 = read_i16_le(params, (i + 2) * 4).unwrap_or(0);
                let y2 = read_i16_le(params, (i + 2) * 4 + 2).unwrap_or(0);
                let x3 = read_i16_le(params, (i + 3) * 4).unwrap_or(0);
                let y3 = read_i16_le(params, (i + 3) * 4 + 2).unwrap_or(0);
                let x4 = read_i16_le(params, (i + 4) * 4).unwrap_or(0);
                let y4 = read_i16_le(params, (i + 4) * 4 + 2).unwrap_or(0);
                let x5 = read_i16_le(params, (i + 5) * 4).unwrap_or(0);
                let y5 = read_i16_le(params, (i + 5) * 4 + 2).unwrap_or(0);
                let x6 = read_i16_le(params, (i + 6) * 4).unwrap_or(0);
                let y6 = read_i16_le(params, (i + 6) * 4 + 2).unwrap_or(0);
                let x7 = read_i16_le(params, (i + 7) * 4).unwrap_or(0);
                let y7 = read_i16_le(params, (i + 7) * 4 + 2).unwrap_or(0);
                let x8 = read_i16_le(params, (i + 8) * 4).unwrap_or(0);
                let y8 = read_i16_le(params, (i + 8) * 4 + 2).unwrap_or(0);
                let x9 = read_i16_le(params, (i + 9) * 4).unwrap_or(0);
                let y9 = read_i16_le(params, (i + 9) * 4 + 2).unwrap_or(0);
                let x10 = read_i16_le(params, (i + 10) * 4).unwrap_or(0);
                let y10 = read_i16_le(params, (i + 10) * 4 + 2).unwrap_or(0);
                let x11 = read_i16_le(params, (i + 11) * 4).unwrap_or(0);
                let y11 = read_i16_le(params, (i + 11) * 4 + 2).unwrap_or(0);
                let x12 = read_i16_le(params, (i + 12) * 4).unwrap_or(0);
                let y12 = read_i16_le(params, (i + 12) * 4 + 2).unwrap_or(0);
                let x13 = read_i16_le(params, (i + 13) * 4).unwrap_or(0);
                let y13 = read_i16_le(params, (i + 13) * 4 + 2).unwrap_or(0);
                let x14 = read_i16_le(params, (i + 14) * 4).unwrap_or(0);
                let y14 = read_i16_le(params, (i + 14) * 4 + 2).unwrap_or(0);
                let x15 = read_i16_le(params, (i + 15) * 4).unwrap_or(0);
                let y15 = read_i16_le(params, (i + 15) * 4 + 2).unwrap_or(0);

                // Pack into SIMD vectors
                let x_vals = _mm256_set_epi16(
                    x15, x14, x13, x12, x11, x10, x9, x8, x7, x6, x5, x4, x3, x2, x1, x0,
                );
                let y_vals = _mm256_set_epi16(
                    y15, y14, y13, y12, y11, y10, y9, y8, y7, y6, y5, y4, y3, y2, y1, y0,
                );

                min_x = _mm256_min_epi16(min_x, x_vals);
                max_x = _mm256_max_epi16(max_x, x_vals);
                min_y = _mm256_min_epi16(min_y, y_vals);
                max_y = _mm256_max_epi16(max_y, y_vals);

                i += 16;
            }

            // Reduce SIMD results
            let mut temp_min_x = [0i16; 16];
            let mut temp_max_x = [0i16; 16];
            let mut temp_min_y = [0i16; 16];
            let mut temp_max_y = [0i16; 16];
            _mm256_storeu_si256(temp_min_x.as_mut_ptr() as *mut __m256i, min_x);
            _mm256_storeu_si256(temp_max_x.as_mut_ptr() as *mut __m256i, max_x);
            _mm256_storeu_si256(temp_min_y.as_mut_ptr() as *mut __m256i, min_y);
            _mm256_storeu_si256(temp_max_y.as_mut_ptr() as *mut __m256i, max_y);

            *left = temp_min_x.iter().copied().min().unwrap_or(*left);
            *right = temp_max_x.iter().copied().max().unwrap_or(*right);
            *top = temp_min_y.iter().copied().min().unwrap_or(*top);
            *bottom = temp_max_y.iter().copied().max().unwrap_or(*bottom);

            // Handle remaining points with SSE4.1
            if i < count {
                Self::update_from_points_sse41(
                    left,
                    top,
                    right,
                    bottom,
                    &params[i * 4..],
                    count - i,
                );
            }
        }
    }

    /// NEON implementation for aarch64 (processes 8 points at a time using scalar loads)
    ///
    /// # Safety
    ///
    /// NEON is always available on aarch64, but this function is still marked unsafe
    /// due to the use of intrinsics.
    #[cfg(target_arch = "aarch64")]
    #[target_feature(enable = "neon")]
    unsafe fn update_from_points_neon(
        left: &mut i16,
        top: &mut i16,
        right: &mut i16,
        bottom: &mut i16,
        params: &[u8],
        count: usize,
    ) {
        // SAFETY: All NEON intrinsic operations are safe on aarch64
        unsafe {
            let mut min_x = vdupq_n_s16(*left);
            let mut min_y = vdupq_n_s16(*top);
            let mut max_x = vdupq_n_s16(*right);
            let mut max_y = vdupq_n_s16(*bottom);

            let mut i = 0;
            // Process 8 points at a time
            while i + 7 < count && (i * 4 + 32) <= params.len() {
                // Load 8 points as individual i16 values
                let x0 = read_i16_le(params, i * 4).unwrap_or(0);
                let y0 = read_i16_le(params, i * 4 + 2).unwrap_or(0);
                let x1 = read_i16_le(params, (i + 1) * 4).unwrap_or(0);
                let y1 = read_i16_le(params, (i + 1) * 4 + 2).unwrap_or(0);
                let x2 = read_i16_le(params, (i + 2) * 4).unwrap_or(0);
                let y2 = read_i16_le(params, (i + 2) * 4 + 2).unwrap_or(0);
                let x3 = read_i16_le(params, (i + 3) * 4).unwrap_or(0);
                let y3 = read_i16_le(params, (i + 3) * 4 + 2).unwrap_or(0);
                let x4 = read_i16_le(params, (i + 4) * 4).unwrap_or(0);
                let y4 = read_i16_le(params, (i + 4) * 4 + 2).unwrap_or(0);
                let x5 = read_i16_le(params, (i + 5) * 4).unwrap_or(0);
                let y5 = read_i16_le(params, (i + 5) * 4 + 2).unwrap_or(0);
                let x6 = read_i16_le(params, (i + 6) * 4).unwrap_or(0);
                let y6 = read_i16_le(params, (i + 6) * 4 + 2).unwrap_or(0);
                let x7 = read_i16_le(params, (i + 7) * 4).unwrap_or(0);
                let y7 = read_i16_le(params, (i + 7) * 4 + 2).unwrap_or(0);

                // Pack into SIMD vectors
                let x_vals = vcombine_s16(
                    vcreate_s16(
                        (x0 as u64)
                            | ((x1 as u64) << 16)
                            | ((x2 as u64) << 32)
                            | ((x3 as u64) << 48),
                    ),
                    vcreate_s16(
                        (x4 as u64)
                            | ((x5 as u64) << 16)
                            | ((x6 as u64) << 32)
                            | ((x7 as u64) << 48),
                    ),
                );
                let y_vals = vcombine_s16(
                    vcreate_s16(
                        (y0 as u64)
                            | ((y1 as u64) << 16)
                            | ((y2 as u64) << 32)
                            | ((y3 as u64) << 48),
                    ),
                    vcreate_s16(
                        (y4 as u64)
                            | ((y5 as u64) << 16)
                            | ((y6 as u64) << 32)
                            | ((y7 as u64) << 48),
                    ),
                );

                min_x = vminq_s16(min_x, x_vals);
                max_x = vmaxq_s16(max_x, x_vals);
                min_y = vminq_s16(min_y, y_vals);
                max_y = vmaxq_s16(max_y, y_vals);

                i += 8;
            }

            // Reduce SIMD results
            let mut temp_min_x = [0i16; 8];
            let mut temp_max_x = [0i16; 8];
            let mut temp_min_y = [0i16; 8];
            let mut temp_max_y = [0i16; 8];
            vst1q_s16(temp_min_x.as_mut_ptr(), min_x);
            vst1q_s16(temp_max_x.as_mut_ptr(), max_x);
            vst1q_s16(temp_min_y.as_mut_ptr(), min_y);
            vst1q_s16(temp_max_y.as_mut_ptr(), max_y);

            *left = temp_min_x.iter().copied().min().unwrap_or(*left);
            *right = temp_max_x.iter().copied().max().unwrap_or(*right);
            *top = temp_min_y.iter().copied().min().unwrap_or(*top);
            *bottom = temp_max_y.iter().copied().max().unwrap_or(*bottom);

            // Handle remaining points with scalar code
            Self::update_from_points_scalar(left, top, right, bottom, &params[i * 4..], count - i);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Helper to create params buffer from points
    fn create_params_from_points(points: &[(i16, i16)]) -> Vec<u8> {
        let mut params = Vec::with_capacity(points.len() * 4);
        for &(x, y) in points {
            params.extend_from_slice(&x.to_le_bytes());
            params.extend_from_slice(&y.to_le_bytes());
        }
        params
    }

    #[test]
    fn test_update_from_points_scalar() {
        let points = vec![(10, 20), (30, 40), (-5, -10), (100, 200)];
        let params = create_params_from_points(&points);

        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;

        BoundsCalculator::update_from_points_scalar(
            &mut left,
            &mut top,
            &mut right,
            &mut bottom,
            &params,
            points.len(),
        );

        assert_eq!(left, -5);
        assert_eq!(top, -10);
        assert_eq!(right, 100);
        assert_eq!(bottom, 200);
    }

    #[test]
    fn test_update_from_points_empty() {
        let mut left = 10i16;
        let mut top = 20i16;
        let mut right = 30i16;
        let mut bottom = 40i16;

        BoundsCalculator::update_from_points(&mut left, &mut top, &mut right, &mut bottom, &[], 0);

        // Should remain unchanged
        assert_eq!(left, 10);
        assert_eq!(top, 20);
        assert_eq!(right, 30);
        assert_eq!(bottom, 40);
    }

    #[test]
    fn test_update_from_points_single() {
        let points = vec![(50, 60)];
        let params = create_params_from_points(&points);

        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;

        BoundsCalculator::update_from_points(
            &mut left,
            &mut top,
            &mut right,
            &mut bottom,
            &params,
            1,
        );

        assert_eq!(left, 50);
        assert_eq!(top, 60);
        assert_eq!(right, 50);
        assert_eq!(bottom, 60);
    }

    #[test]
    fn test_update_from_points_various_sizes() {
        // Test different sizes to trigger different SIMD code paths
        for count in [2, 3, 4, 5, 7, 8, 9, 15, 16, 17, 31, 32, 33, 64, 100] {
            let points: Vec<(i16, i16)> = (0..count)
                .map(|i| {
                    let x = (i as i16 * 13 + 7) % 1000 - 500;
                    let y = (i as i16 * 17 + 11) % 1000 - 500;
                    (x, y)
                })
                .collect();

            let params = create_params_from_points(&points);

            // Calculate using SIMD
            let mut left_simd = i16::MAX;
            let mut top_simd = i16::MAX;
            let mut right_simd = i16::MIN;
            let mut bottom_simd = i16::MIN;

            BoundsCalculator::update_from_points(
                &mut left_simd,
                &mut top_simd,
                &mut right_simd,
                &mut bottom_simd,
                &params,
                count,
            );

            // Calculate using scalar
            let mut left_scalar = i16::MAX;
            let mut top_scalar = i16::MAX;
            let mut right_scalar = i16::MIN;
            let mut bottom_scalar = i16::MIN;

            BoundsCalculator::update_from_points_scalar(
                &mut left_scalar,
                &mut top_scalar,
                &mut right_scalar,
                &mut bottom_scalar,
                &params,
                count,
            );

            assert_eq!(left_simd, left_scalar, "Left mismatch for count {}", count);
            assert_eq!(top_simd, top_scalar, "Top mismatch for count {}", count);
            assert_eq!(
                right_simd, right_scalar,
                "Right mismatch for count {}",
                count
            );
            assert_eq!(
                bottom_simd, bottom_scalar,
                "Bottom mismatch for count {}",
                count
            );
        }
    }

    #[test]
    fn test_update_from_points_negative_coords() {
        let points = vec![(-100, -200), (-50, -75), (0, 0), (50, 75)];
        let params = create_params_from_points(&points);

        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;

        BoundsCalculator::update_from_points(
            &mut left,
            &mut top,
            &mut right,
            &mut bottom,
            &params,
            points.len(),
        );

        assert_eq!(left, -100);
        assert_eq!(top, -200);
        assert_eq!(right, 50);
        assert_eq!(bottom, 75);
    }

    #[test]
    fn test_update_from_points_extreme_values() {
        let points = vec![(i16::MIN, i16::MIN), (i16::MAX, i16::MAX), (0, 0), (-1, 1)];
        let params = create_params_from_points(&points);

        let mut left = i16::MAX;
        let mut top = i16::MAX;
        let mut right = i16::MIN;
        let mut bottom = i16::MIN;

        BoundsCalculator::update_from_points(
            &mut left,
            &mut top,
            &mut right,
            &mut bottom,
            &params,
            points.len(),
        );

        assert_eq!(left, i16::MIN);
        assert_eq!(top, i16::MIN);
        assert_eq!(right, i16::MAX);
        assert_eq!(bottom, i16::MAX);
    }

    #[test]
    fn test_update_from_points_with_initial_bounds() {
        let points = vec![(20, 30), (40, 50)];
        let params = create_params_from_points(&points);

        let mut left = 10i16;
        let mut top = 15i16;
        let mut right = 60i16;
        let mut bottom = 70i16;

        BoundsCalculator::update_from_points(
            &mut left,
            &mut top,
            &mut right,
            &mut bottom,
            &params,
            points.len(),
        );

        // Should keep tighter bounds
        assert_eq!(left, 10); // Kept initial smaller value
        assert_eq!(top, 15); // Kept initial smaller value
        assert_eq!(right, 60); // Kept initial larger value
        assert_eq!(bottom, 70); // Kept initial larger value
    }

    /// Test direct SIMD implementations against scalar
    #[cfg(target_arch = "x86_64")]
    #[test]
    fn test_x86_sse2_vs_scalar() {
        let points: Vec<(i16, i16)> = (0..64)
            .map(|i| ((i * 13 - 100) as i16, (i * 17 - 200) as i16))
            .collect();
        let params = create_params_from_points(&points);

        let mut left_sse2 = i16::MAX;
        let mut top_sse2 = i16::MAX;
        let mut right_sse2 = i16::MIN;
        let mut bottom_sse2 = i16::MIN;

        unsafe {
            BoundsCalculator::update_from_points_sse2(
                &mut left_sse2,
                &mut top_sse2,
                &mut right_sse2,
                &mut bottom_sse2,
                &params,
                points.len(),
            );
        }

        let mut left_scalar = i16::MAX;
        let mut top_scalar = i16::MAX;
        let mut right_scalar = i16::MIN;
        let mut bottom_scalar = i16::MIN;

        BoundsCalculator::update_from_points_scalar(
            &mut left_scalar,
            &mut top_scalar,
            &mut right_scalar,
            &mut bottom_scalar,
            &params,
            points.len(),
        );

        assert_eq!(left_sse2, left_scalar);
        assert_eq!(top_sse2, top_scalar);
        assert_eq!(right_sse2, right_scalar);
        assert_eq!(bottom_sse2, bottom_scalar);
    }

    #[cfg(target_arch = "x86_64")]
    #[test]
    fn test_x86_sse41_vs_scalar() {
        if !is_x86_feature_detected!("sse4.1") {
            eprintln!("SSE4.1 not available, skipping test");
            return;
        }

        let points: Vec<(i16, i16)> = (0..64)
            .map(|i| ((i * 13 - 100) as i16, (i * 17 - 200) as i16))
            .collect();
        let params = create_params_from_points(&points);

        let mut left_sse41 = i16::MAX;
        let mut top_sse41 = i16::MAX;
        let mut right_sse41 = i16::MIN;
        let mut bottom_sse41 = i16::MIN;

        unsafe {
            BoundsCalculator::update_from_points_sse41(
                &mut left_sse41,
                &mut top_sse41,
                &mut right_sse41,
                &mut bottom_sse41,
                &params,
                points.len(),
            );
        }

        let mut left_scalar = i16::MAX;
        let mut top_scalar = i16::MAX;
        let mut right_scalar = i16::MIN;
        let mut bottom_scalar = i16::MIN;

        BoundsCalculator::update_from_points_scalar(
            &mut left_scalar,
            &mut top_scalar,
            &mut right_scalar,
            &mut bottom_scalar,
            &params,
            points.len(),
        );

        assert_eq!(left_sse41, left_scalar);
        assert_eq!(top_sse41, top_scalar);
        assert_eq!(right_sse41, right_scalar);
        assert_eq!(bottom_sse41, bottom_scalar);
    }

    #[cfg(target_arch = "x86_64")]
    #[test]
    fn test_x86_avx2_vs_scalar() {
        if !is_x86_feature_detected!("avx2") {
            eprintln!("AVX2 not available, skipping test");
            return;
        }

        let points: Vec<(i16, i16)> = (0..128)
            .map(|i| ((i * 13 - 100) as i16, (i * 17 - 200) as i16))
            .collect();
        let params = create_params_from_points(&points);

        let mut left_avx2 = i16::MAX;
        let mut top_avx2 = i16::MAX;
        let mut right_avx2 = i16::MIN;
        let mut bottom_avx2 = i16::MIN;

        unsafe {
            BoundsCalculator::update_from_points_avx2(
                &mut left_avx2,
                &mut top_avx2,
                &mut right_avx2,
                &mut bottom_avx2,
                &params,
                points.len(),
            );
        }

        let mut left_scalar = i16::MAX;
        let mut top_scalar = i16::MAX;
        let mut right_scalar = i16::MIN;
        let mut bottom_scalar = i16::MIN;

        BoundsCalculator::update_from_points_scalar(
            &mut left_scalar,
            &mut top_scalar,
            &mut right_scalar,
            &mut bottom_scalar,
            &params,
            points.len(),
        );

        assert_eq!(left_avx2, left_scalar);
        assert_eq!(top_avx2, top_scalar);
        assert_eq!(right_avx2, right_scalar);
        assert_eq!(bottom_avx2, bottom_scalar);
    }

    /// Test scan_records with actual WMF record structures
    #[test]
    fn test_scan_records_polygon() {
        use bytes::Bytes;

        let points = vec![(10, 20), (30, 40), (-5, -10), (100, 200)];
        let mut params = Vec::new();
        params.extend_from_slice(&(points.len() as i16).to_le_bytes());
        params.extend_from_slice(&create_params_from_points(&points));

        let records = vec![WmfRecord {
            size: 0,
            function: record::POLYGON,
            params: Bytes::from(params),
        }];

        let (left, top, right, bottom) = BoundsCalculator::scan_records(&records);

        assert_eq!(left, -5);
        assert_eq!(top, -10);
        assert_eq!(right, 100);
        assert_eq!(bottom, 200);
    }

    #[test]
    fn test_scan_records_multiple_shapes() {
        use bytes::Bytes;

        // Rectangle record
        let rect_params = {
            let mut p = Vec::new();
            p.extend_from_slice(&50i16.to_le_bytes()); // bottom
            p.extend_from_slice(&150i16.to_le_bytes()); // right
            p.extend_from_slice(&10i16.to_le_bytes()); // top
            p.extend_from_slice(&50i16.to_le_bytes()); // left
            p
        };

        // Polygon record
        let polygon_points = vec![(200, 300), (250, 350)];
        let mut polygon_params = Vec::new();
        polygon_params.extend_from_slice(&(polygon_points.len() as i16).to_le_bytes());
        polygon_params.extend_from_slice(&create_params_from_points(&polygon_points));

        let records = vec![
            WmfRecord {
                size: 0,
                function: record::RECTANGLE,
                params: Bytes::from(rect_params),
            },
            WmfRecord {
                size: 0,
                function: record::POLYGON,
                params: Bytes::from(polygon_params),
            },
        ];

        let (left, top, right, bottom) = BoundsCalculator::scan_records(&records);

        assert_eq!(left, 50);
        assert_eq!(top, 10);
        assert_eq!(right, 250);
        assert_eq!(bottom, 350);
    }

    #[test]
    fn test_scan_records_empty() {
        let records = vec![];
        let (left, top, right, bottom) = BoundsCalculator::scan_records(&records);

        // Should return default bounds
        assert_eq!(left, 0);
        assert_eq!(top, 0);
        assert_eq!(right, 1000);
        assert_eq!(bottom, 1000);
    }
}
