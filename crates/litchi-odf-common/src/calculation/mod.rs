//! Shared ODF formula-calculation settings.
//!
//! The focused module owns the schema model and the bounded XML codec used by
//! spreadsheet-like ODF families, including chart documents that carry the
//! same inert metadata.  Family crates only compose these values into their
//! package and facade layers.

mod codec;
mod model;
#[cfg(test)]
mod tests;

pub use codec::{parse, write};
pub use model::{Iteration, IterationStatus, NullDate, Settings};

// Fused-parser plumbing — kept `pub` (not `pub(crate)`) because family owner
// crates drive the handler across the crate boundary. `#[doc(hidden)]`
// suppresses the public docs surface.
#[doc(hidden)]
pub use codec::{CalculationHandler, classify, validate_size};

pub(crate) const MAX_XML_BYTES: usize = 64 * 1_048_576;
pub(crate) const MAX_DEPTH: usize = 256;
pub(crate) const MAX_EVENTS: usize = 1_000_000;
pub(crate) const MAX_ATTRIBUTE_BYTES: usize = 64 * 1_024;
