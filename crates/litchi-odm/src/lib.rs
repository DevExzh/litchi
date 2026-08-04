//! OpenDocument Master Document support with semantic responsibility layers.
#![forbid(unsafe_code)]

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod model;
pub mod package;

pub use facade::{Builder, Master};
