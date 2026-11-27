//! Bounded authoring of the notes-master text-style round-trip package.
//!
//! `[MS-PPT]` §2.11.18 stores an ECMA-376 package in a
//! `RoundTripNotesMasterTextStyles12Atom`.  The package is intentionally kept
//! behind this contextual owner: callers work with `Styles`, while the
//! package and record framing remain validated implementation details.

pub(crate) mod codec;
pub(crate) mod package;
pub(crate) mod validation;

mod model;

#[cfg(test)]
mod tests;

pub use model::{MAX_PACKAGE_BYTES, MAX_PARTS, MAX_XML_BYTES, Package, Styles};
