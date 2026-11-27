//! Neutral, bounded BIFF record framing.
//!
//! This crate implements only the physical record substrate shared by BIFF
//! consumers. A frame contains a little-endian two-byte [`Kind`], a little-
//! endian two-byte payload length, and exactly that many payload bytes. It
//! deliberately does not interpret record kinds, continuation records, chart
//! grammar, workbook topology, or any other host-specific semantics.
//!
//! The wire shape follows `[MS-XLS]` section 2.1.4 and `[MS-OGRAPH]` section
//! 2.1.4. Unknown kinds and payload bytes are accepted and retained so that a
//! higher-level owner can preserve records it does not understand.

#![forbid(unsafe_code)]

pub mod error;
pub mod frame;
pub mod limits;
pub mod stream;

pub use error::{Error, Resource, Result};
pub use frame::{Kind, Record, RecordRef};
pub use limits::{Limits, MAX_RECORD_BYTES};
pub use stream::{Encoder, Records};
