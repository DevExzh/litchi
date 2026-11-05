//! Layered `PresentationML` timing codec facade.
//!
//! [`semantic`] owns typed model operations, [`xml`] owns namespace-aware
//! parsing/writing, and [`validation`] owns bounded values and safety policy.

mod semantic;
mod validation;
mod xml;

pub use validation::MAX_TIMING_MILLISECONDS;
pub(super) use validation::{
    MAX_NORMALIZED_TIME_DECIMALS, MAX_TEMPLATE_TIME_NODE_BYTES, MAX_TIME_FILTER_BYTES,
    MAX_TIME_FILTER_POINTS, MAX_TIMING_ATTRIBUTES, MAX_TIMING_DEPTH, MAX_TIMING_NODES,
    MAX_TIMING_TEXT_BYTES, MAX_TIMING_XML_BYTES,
};

#[cfg(test)]
pub(super) use validation::MAX_ANIMATION_BUILDS;
