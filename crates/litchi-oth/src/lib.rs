//! OpenDocument HTML Template support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{Builder, Template};
pub use model::{link, paragraph};
