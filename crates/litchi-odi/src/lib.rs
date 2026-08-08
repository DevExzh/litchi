//! `OpenDocument` Image support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod flat;
mod model;
mod package;

pub use facade::{Builder, Image};
pub use flat::FlatImage;
pub use model::{frame, source};
