//! Common SIMD operations
//!
//! This module provides high-performance SIMD (Single Instruction, Multiple Data) operations
//! optimized for various CPU architectures and instruction sets.
//!
//! # Supported Architectures
//!
//! ## x86_64
//! - **SSE** (Streaming SIMD Extensions): 128-bit vectors
//! - **SSE2**: Enhanced 128-bit integer operations
//! - **SSE3**: Additional 128-bit operations
//! - **SSSE3**: Supplemental 128-bit operations
//! - **SSE4.1**: 128-bit operations with additional instructions
//! - **SSE4.2**: 128-bit operations with string/text processing
//! - **AVX** (Advanced Vector Extensions): 256-bit floating-point operations
//! - **AVX2**: 256-bit integer operations
//! - **AVX-512**: 512-bit operations (F, BW, DQ, VL extensions)
//!
//! ## aarch64 (ARM)
//! - **NEON**: 128-bit SIMD operations
//! - **SVE** (Scalable Vector Extension): Variable-length vectors (future support)
//! - **SVE2**: Enhanced SVE operations (future support)
//!
//! # Modules
//!
//! - [`cmp`]: Vector comparison operations (equal, not equal, greater than, less than, etc.)
//!
//! # Performance Considerations
//!
//! This module is designed with performance as the top priority:
//!
//! - **Runtime Feature Detection**: Automatically selects the best available instruction set
//! - **Zero-Copy Operations**: Leverages Rust's ownership system to avoid unnecessary allocations
//! - **Inline Functions**: All hot-path functions are marked `#[inline]` for optimal performance
//! - **Cache-Friendly**: Operations are designed to maximize CPU cache utilization
//! - **Minimal Overhead**: Direct mapping to hardware instructions where possible
//!
//! # Examples
//!
//! ```rust
//! use litchi::common::simd::cmp::simd_eq_u8;
//!
//! // Compare two byte arrays for equality
//! let a = vec![1u8, 2, 3, 4, 5, 6, 7, 8];
//! let b = vec![1u8, 2, 0, 4, 5, 0, 7, 8];
//! let mut result = vec![0u8; 8];
//!
//! simd_eq_u8(&a, &b, &mut result);
//! // result[i] is 0xFF where a[i] == b[i], 0x00 otherwise
//! ```
//!
//! # Safety
//!
//! Functions using SIMD intrinsics are marked as `unsafe` when they require specific CPU features.
//! High-level API functions perform runtime feature detection to ensure safety across different CPUs.
//!
//! When using low-level intrinsics directly, ensure the target CPU supports the required features
//! either through:
//! - Runtime detection with `is_x86_feature_detected!()` or similar
//! - Compile-time target features: `#[target_feature(enable = "avx2")]`
//! - Compiler flags: `RUSTFLAGS="-C target-feature=+avx2"`

pub mod cmp;
