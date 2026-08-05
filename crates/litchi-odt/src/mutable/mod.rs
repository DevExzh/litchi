//! Mutable document structure for in-place modifications.
//!
//! The owner keeps mutable document state and element operations in `model`,
//! content and styles XML snapshots and edits in `codec`, package input and
//! output in `package`, and regression coverage in `tests`.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::MutableDocument;
