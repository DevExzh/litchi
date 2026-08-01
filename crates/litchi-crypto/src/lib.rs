//! Format-neutral Microsoft Office cryptography structures.
//!
//! This crate owns bounded, runtime-independent `[MS-OFFCRYPTO]` parsing and
//! transformation primitives. Concrete document formats remain responsible
//! for locating records and mapping typed failures into their own errors.

#![forbid(unsafe_code)]

pub mod integrity;
pub mod labels;
pub mod protected;
pub mod rc4;
pub mod spaces;
