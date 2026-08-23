//! Validated low-level BIFF12 records and values.
//!
//! The slice-backed path lends record payloads from the caller's allocation.
//! No parser in this module accepts an unbounded payload or string.

mod cursor;
mod error;
pub mod kind;
mod record;
mod writer;

pub use cursor::Cursor;
pub use error::{Error, LimitResource, Result, Stage};
pub use record::{Header, Kind, Limits, MAX_WIRE_PAYLOAD, MAX_WIRE_STRING_UNITS, Record, Records};
pub use writer::Writer;
