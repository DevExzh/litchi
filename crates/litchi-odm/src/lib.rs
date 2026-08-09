//! `OpenDocument` Master Document support with semantic responsibility layers.
#![forbid(unsafe_code)]

mod authoring;
mod codec;
mod facade;
pub mod link;
mod model;
mod package;
pub mod title;

pub use facade::{Builder, Master};
pub use model::{section, subdocument};
