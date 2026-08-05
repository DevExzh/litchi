//! Layered PresentationML timing codec facade.
//!
//! [`semantic`] owns typed model operations, [`xml`] owns namespace-aware
//! parsing/writing, and [`validation`] owns bounded values and safety policy.

mod semantic;
mod validation;
mod xml;

pub use validation::MAX_TIMING_MILLISECONDS;
pub(super) use validation::{
    MAX_NORMALIZED_TIME_DECIMALS, MAX_TIME_FILTER_BYTES, MAX_TIME_FILTER_POINTS,
};

#[cfg(test)]
pub(super) use validation::{MAX_ANIMATION_BUILDS, MAX_TIMING_DEPTH, MAX_TIMING_NODES};
