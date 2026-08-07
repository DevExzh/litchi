//! `OpenDocument` Drawing support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{Builder, Drawing};
pub use model::{layer, page};
