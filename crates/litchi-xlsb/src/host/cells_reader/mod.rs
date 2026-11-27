//! Layered BIFF12 worksheet cell reader facade.
//!
//! The facade keeps the package-facing reader compact while separating typed
//! state, stream decoding, semantic conversions, and record validation.

mod codec;
mod model;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) use model::CellsReader;
#[allow(
    unused_imports,
    reason = "module re-exports preserve stable wire and semantic API paths"
)]
pub(crate) use model::Dimensions;
