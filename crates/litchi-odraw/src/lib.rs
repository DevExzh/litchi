//! Format-neutral Microsoft OfficeArt drawing primitives.
//!
//! The crate provides checked, zero-copy parsing over caller-owned bytes and
//! runtime-neutral writing helpers shared by the legacy Office formats.

#![forbid(unsafe_code)]
#![allow(
    non_ascii_idents,
    reason = "zerocopy's wire-layout derives emit compiler-generated identifiers"
)]

mod container;
mod error;
pub mod image;
mod parser;
pub mod prop;
mod record;
pub mod rule;
pub mod shape;
pub mod write;

pub use container::{Children, Container, Limits};
pub use error::{Error, ImageLimit, Limit, Result};
pub use parser::Parser;
pub use record::{Record, RecordKind};
