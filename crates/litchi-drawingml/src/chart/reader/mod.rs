//! DrawingML chart XML reader facade.
//!
//! The public reader path remains [`crate::chart::reader`]. Namespace-aware
//! stream state lives in `model`, XML conversion lives in `codec`, and the
//! focused conformance suite stays in `tests`.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub use codec::read;
