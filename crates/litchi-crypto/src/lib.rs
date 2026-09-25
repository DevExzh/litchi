//! Format-neutral Microsoft Office cryptography structures.
//!
//! This crate owns bounded, runtime-independent `[MS-OFFCRYPTO]` parsing and
//! transformation primitives. Concrete document formats remain responsible
//! for locating records and mapping typed failures into their own errors.

#![forbid(unsafe_code)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `BytesStartExt` (record 0770, workspace `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

pub mod integrity;
pub mod labels;
pub mod legacy_rc4;
#[cfg(feature = "ooxml")]
pub mod ooxml;
pub mod protected;
pub mod rc4;
pub mod spaces;
