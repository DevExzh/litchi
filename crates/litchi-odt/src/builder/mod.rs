//! `OpenDocument` Text document builder.
//!
//! This module provides a builder pattern for creating new ODT documents from scratch.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use model::Builder;
