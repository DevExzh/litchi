//! WordprocessingML package owner.
//!
//! The model stores typed package state, the codec owns OPC I/O, and the
//! package layer coordinates relationship-backed graph edits.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::Package;
