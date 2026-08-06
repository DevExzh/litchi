//! Chart aggregate facade.
//!
//! The public [`Chart`] model stays intentionally small at this boundary;
//! semantic mutation lives in [`semantic`], while derived cache invariants
//! live in [`validation`].

mod semantic;
mod validation;

#[cfg(test)]
mod tests;

pub use semantic::Chart;

pub(in crate::chart) use validation::{cache_dimensions, dimensions_cover};
