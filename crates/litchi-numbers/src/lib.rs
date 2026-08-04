//! Numbers semantic value models.
//!
//! Archive parsing, protobuf decoding, and package mutation remain owned by
//! the Numbers implementation. This crate starts the downward migration with
//! the dependency-free cell vocabulary used by Numbers, Pages table editing,
//! and the shared structured extractor.

#![forbid(unsafe_code)]

/// Cell-level Numbers vocabulary.
pub mod cell;
