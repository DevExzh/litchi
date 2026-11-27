//! Host-neutral `SpreadsheetDrawing` shapes.
//!
//! `read` is lossless for supported inspection payloads, retaining unknown
//! attributes and elements. `writer` authors fresh objects only; it does not
//! serialize opaque reader payloads.

mod model;
mod reader;
pub mod writer;

pub use model::*;
pub use reader::read;

#[cfg(test)]
mod tests;
