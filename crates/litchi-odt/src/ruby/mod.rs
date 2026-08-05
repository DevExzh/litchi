//! Semantic parsing of OpenDocument ruby annotations.

mod codec;
mod model;

#[cfg(test)]
mod tests;

const MAX_DEPTH: usize = 4_096;
const MAX_RUBIES: usize = 1_000_000;

pub use model::Annotation;

pub(crate) use codec::parse_rubies;
