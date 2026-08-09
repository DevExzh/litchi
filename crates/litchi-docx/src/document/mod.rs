//! Semantic WordprocessingML main-document facade.
//!
//! The document owner keeps its public model, document-XML codec, and
//! package-bound orchestration in separate layers while retaining the
//! historical `crate::document` entry point.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

pub use model::{Block, Document, Element, ImageWatermarkPart, OpaqueBlock};
