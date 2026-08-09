//! Exact, checked `DrawingML` coordinate domains.
//!
//! [`Coordinate`] models the complete `a:ST_Coordinate` union: either a
//! bounded unqualified EMU integer or an exact `s:ST_UniversalMeasure` value.
//! [`Extent`] models the integer-only `a:ST_PositiveCoordinate` restriction.
//!
//! The semantic values live in `model`, lexical construction and formatting
//! live in `codec`, and the focused conformance checks live in `tests`.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::ParseError;
pub use model::{Coordinate, Extent, MAX_BYTES, MAX_EMU, MIN_EMU, Unit};
