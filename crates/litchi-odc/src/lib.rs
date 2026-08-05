//! OpenDocument Chart support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{Builder, Chart};
pub use model::{axis, legend, series};
