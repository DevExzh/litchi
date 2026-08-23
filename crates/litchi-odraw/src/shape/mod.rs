//! Typed, bounded views over `OfficeArt` shape containers.
//!
//! The facade keeps the ergonomic shape API flat while the implementation is
//! split by responsibility: `model` owns borrowed typed objects, `codec`
//! owns wire traversal, and `validation` owns topology and record checks.

mod codec;
mod model;
mod validation;

pub use codec::{parse, parse_with};
pub use model::{Bounds, Flags, Kind, Native, Shape};

#[cfg(test)]
mod tests;
