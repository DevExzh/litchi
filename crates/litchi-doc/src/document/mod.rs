//! Typed Word document state, package orchestration, and binary codecs.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::Document;
