//! Typed, lossless MS-DOC routing-slip metadata.
//!
//! The facade exposes byte-oriented ANSI fields so parsing never assumes a
//! UTF-8 code page or changes the bytes that Word stored. [`Metadata::parse`]
//! reads the optional table range addressed by FIB index 70; [`parse_bytes`]
//! and [`to_bytes`] operate on one complete `Metadata` payload.
//!
//! The context intentionally stops at the FIB/table-stream seam. `Document`
//! and its package owners do not currently expose routing slips, so callers
//! that need this metadata supply the already-selected table stream and its
//! [`crate::parts::fib::FileInformationBlock`]. This keeps the feature isolated and avoids
//! changing unrelated document ownership or save paths.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::{parse, parse_bytes, to_bytes};
pub use model::{DeliveryOption, Metadata, NarrowString, Protection, Recipient};
