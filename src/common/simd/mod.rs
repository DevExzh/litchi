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
//! - **NEON**: Fixed 128-bit SIMD operations (always available on aarch64)
//! - **SVE** (Scalable Vector Extension): Variable-length vectors (128-2048 bits)
//! - **SVE2**: Enhanced SVE with additional DSP and multimedia operations
//!
//! # Modules
//!
//! - [`cmp`]: Vector comparison operations (equal, not equal, greater than, less than, etc.)
//! - [`fmt`]: SIMD-optimized formatting operations (hex encoding, GUID/CLSID formatting)
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
//! # SVE/SVE2 Features
//!
//! ARM's Scalable Vector Extension (SVE) and SVE2 provide unique capabilities:
//!
//! ## SVE Key Features
//!
//! - **Scalable**: Vector length determined at runtime (128-2048 bits in 128-bit increments)
//! - **Predicated**: All operations use predicate masks for efficient conditional execution
//! - **Loop-friendly**: `svwhilelt` and similar instructions simplify vectorized loops
//! - **Future-proof**: Same code automatically leverages larger vectors on newer hardware
//!
//! ## SVE2 Additions
//!
//! SVE2 extends SVE with instructions for:
//! - DSP operations (saturating arithmetic, complex numbers)
//! - Multimedia processing (polynomial math, CRC)
//! - Bit manipulation (population count, bit permutations)
//! - Table operations and histogram processing
//!
//! ## Availability
//!
//! - **SVE**: Available on some ARMv8.2-A+ processors (e.g., Fujitsu A64FX, AWS Graviton3)
//! - **SVE2**: Available on ARMv9-A processors (e.g., Apple M4, AWS Graviton4)
//! - **Detection**: Compile with `+sve` or `+sve2` target features, or check HWCAP at runtime
//!
//! # Safety
//!
//! Functions using SIMD intrinsics are marked as `unsafe` when they require specific CPU features.
//! High-level API functions perform runtime feature detection to ensure safety across different CPUs.
//!
//! When using low-level intrinsics directly, ensure the target CPU supports the required features
//! either through:
//! - Runtime detection with `is_x86_feature_detected!()` on x86_64
//! - Compile-time target features: `#[target_feature(enable = "avx2")]` or `#[target_feature(enable = "sve")]`
//! - Compiler flags: `RUSTFLAGS="-C target-feature=+avx2"` or `RUSTFLAGS="-C target-feature=+sve"`
//!
//! ## Example: Compiling with SVE Support
//!
//! ```bash
//! # For SVE
//! RUSTFLAGS="-C target-cpu=native -C target-feature=+sve" cargo build --release
//!
//! # For SVE2
//! RUSTFLAGS="-C target-cpu=native -C target-feature=+sve2" cargo build --release
//! ```

pub mod cmp;
pub mod fmt;
