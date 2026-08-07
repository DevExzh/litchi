//! `OpenDocument` Master Document support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
mod model;
mod package;

pub use facade::{Builder, Master};
pub use model::{section, subdocument};
